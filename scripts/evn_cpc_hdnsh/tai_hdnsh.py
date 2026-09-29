"""Tải hợp đồng mua bán điện ngoài sinh hoạt (mã hồ sơ HDNSH) từ cskh.cpc.vn theo danh sách Mã KH EVN.

Luồng xử lý (mỗi Mã KH):
	1. Mở trang Quản lý hợp đồng, nhập Mã khách hàng, bấm Tìm kiếm.
	2. Lọc các dòng có MÃ HỒ SƠ = HDNSH (có thể nhiều dòng / khách hàng).
	3. Với từng dòng: bấm "Xem chi tiết", bắt file PDF (download / popup / iframe / blob / base64 trong JSON) và lưu vào
	   <output>/<Site_id>_<MaKH>/<Site_id>_<MaKH>_HDNSH_<n>.pdf
	4. Xác minh: Mã KH có xuất hiện trong nội dung PDF (nếu PDF có lớp text), đánh dấu file trùng nội dung (SHA-256).
	5. Ghi kết quả vào progress.csv (chạy lại sẽ tiếp tục mã chưa xong) và báo cáo Excel tổng hợp.

Cài đặt (Windows, 1 lần):
	pip install playwright pandas openpyxl python-dotenv pypdf
	(tùy chọn) pip install keyring        # đọc mật khẩu từ Windows Credential Manager
	(tùy chọn) playwright install chromium # nếu không dùng Microsoft Edge có sẵn trên máy

Cấu hình: tạo file .env cạnh script (xem .env.example). KHÔNG ghi mật khẩu vào code.

Chạy thử 3 mã trước, kiểm tra kết quả rồi mới chạy toàn bộ:
	python tai_hdnsh.py --mapping mapping_MaKH_SiteID_GiaLai.xlsx --limit 3
	python tai_hdnsh.py --mapping mapping_MaKH_SiteID_GiaLai.xlsx
	python tai_hdnsh.py --mapping ... --only PC07AA0144953,PC10AA0123456   # chạy lại vài mã cụ thể
"""

from __future__ import annotations

import argparse
import asyncio
import base64
import csv
import getpass
import hashlib
import json
import os
import random
import re
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote, urljoin

import pandas as pd
from dotenv import load_dotenv
from playwright.async_api import BrowserContext, Download, Frame, Locator, Page, Request, Response, async_playwright
from playwright.async_api import TimeoutError as PWTimeout

MA_HO_SO = 'HDNSH'
CONTRACT_PATH = '/dashboard/quan-ly-hop-dong'
MA_KH_RE = re.compile(r'^P[A-Z]\d{2}[A-Z]{2}\d{7}$')
DETAIL_TEXT = re.compile(r'xem\s*chi\s*ti[eế]t', re.I)
DOWNLOAD_TEXT = re.compile(r't[aả]i\s*(xu[oố]ng|v[eề])|download|t[aả]i\s*file', re.I)
CLOSE_TEXT = re.compile(r'^\s*(đóng|dong|close|quay lại|quay lai|×|✕)\s*$', re.I)
LOGIN_TEXT = re.compile(r'đăng\s*nhập|dang\s*nhap|login|sign\s*in', re.I)
EMPTY_TEXT = re.compile(r'không\s*có\s*dữ\s*liệu|no\s*data|trống', re.I)
PDF_B64_RE = re.compile(rb'"(JVBERi0[A-Za-z0-9+/=\\rn]{200,})"')  # base64 của "%PDF-" nằm trong JSON
DIALOG_SEL = (
	'[role=dialog]:visible, .modal.show, .ant-modal-wrap:visible, .ant-drawer-open, .MuiDialog-root, '
	'.el-dialog__wrapper:visible, .v-dialog--active, .p-dialog:visible'
)

PROGRESS_FIELDS = [
	'Site_id',
	'Mã KH EVN',
	'Mã trạm biến áp',
	'Số dòng HDNSH trên web',
	'Số file tải được',
	'Trạng thái',
	'Mã KH có trong PDF',
	'File trùng nội dung',
	'Thông tin dòng HDNSH trên web',
	'Thư mục',
	'Ghi chú',
	'Thời gian',
]
DONE_STATUSES = ('OK', 'KHÔNG CÓ HDNSH', 'CẦN KIỂM TRA')


@dataclass
class Customer:
	site_id: str
	ma_kh: str
	ma_tba: str = ''

	@property
	def folder_name(self) -> str:
		return safe_name(f'{self.site_id}_{self.ma_kh}' if self.site_id else self.ma_kh)


@dataclass
class RowInfo:
	index: int
	text: str


@dataclass
class CustomerResult:
	customer: Customer
	rows: list[RowInfo] = field(default_factory=list)
	files: list[Path] = field(default_factory=list)
	status: str = ''
	verify: str = ''
	duplicates: str = ''
	notes: list[str] = field(default_factory=list)


# ----------------------------------------------------------------------------- tiện ích


def safe_name(s: str) -> str:
	return re.sub(r'[\\/:*?"<>|\s]+', '_', s.strip()).strip('_') or 'khong_ten'


def sha256(path: Path) -> str:
	return hashlib.sha256(path.read_bytes()).hexdigest()


def ext_for(content_type: str, data: bytes) -> str:
	if data[:5] == b'%PDF-':
		return '.pdf'
	if data[:3] == b'\xff\xd8\xff':
		return '.jpg'
	if data[:8] == b'\x89PNG\r\n\x1a\n':
		return '.png'
	ct = content_type.lower()
	for key, ext in (('pdf', '.pdf'), ('jpeg', '.jpg'), ('png', '.png'), ('zip', '.zip'), ('word', '.docx')):
		if key in ct:
			return ext
	return '.zip' if data[:2] == b'PK' else '.bin'


def pdf_contains(path: Path, needle: str) -> bool | None:
	"""True/False nếu đọc được text của PDF, None nếu PDF dạng ảnh scan / không đọc được."""
	if path.suffix.lower() != '.pdf':
		return None
	try:
		from pypdf import PdfReader
	except ImportError:
		return None
	try:
		text = ''.join((page.extract_text() or '') for page in PdfReader(str(path)).pages)
	except Exception:
		return None
	compact = re.sub(r'\s+', '', text).upper()
	if len(compact) < 20:
		return None
	return needle.upper() in compact


def load_customers(mapping: Path, chi_nhanh: str, only: set[str]) -> list[Customer]:
	"""Đọc mapping (xuất từ Google Sheet 'EVN data' -> tab 'DS mã KH EVN', dạng .xlsx hoặc .csv)."""
	if mapping.suffix.lower() in ('.xlsx', '.xlsm', '.xls'):
		xl = pd.ExcelFile(mapping)
		sheet = 'DS mã KH EVN' if 'DS mã KH EVN' in xl.sheet_names else xl.sheet_names[0]
		df = xl.parse(sheet, dtype=str)
	else:
		df = pd.read_csv(mapping, dtype=str, encoding='utf-8-sig')
	df = df.fillna('')
	df.columns = [str(c).strip() for c in df.columns]
	for col in df.columns:
		df[col] = df[col].astype(str).str.strip()
	assert 'Mã KH EVN' in df.columns, f'Không thấy cột "Mã KH EVN" trong {mapping}. Các cột: {list(df.columns)}'
	if chi_nhanh and 'Chi nhánh' in df.columns:
		df = df[df['Chi nhánh'].str.casefold() == chi_nhanh.casefold()]
	customers: list[Customer] = []
	seen: set[str] = set()
	for _, r in df.iterrows():
		ma_kh = r['Mã KH EVN'].upper().replace(' ', '')
		if not ma_kh or ma_kh in seen or (only and ma_kh not in only):
			continue
		if not MA_KH_RE.match(ma_kh):
			print(f'  [BỎ QUA] Mã KH sai định dạng: {ma_kh!r} (Site {r.get("Site_id", "")})')
			continue
		seen.add(ma_kh)
		customers.append(Customer(site_id=r.get('Site_id', ''), ma_kh=ma_kh, ma_tba=r.get('Mã trạm biến áp', '')))
	return customers


def load_done(progress_csv: Path) -> set[str]:
	if not progress_csv.exists():
		return set()
	latest: dict[str, str] = {}
	with progress_csv.open(encoding='utf-8-sig', newline='') as f:
		for row in csv.DictReader(f):
			latest[row['Mã KH EVN']] = row.get('Trạng thái', '')
	return {k for k, v in latest.items() if v.startswith(DONE_STATUSES)}


def append_progress(progress_csv: Path, res: CustomerResult, out_dir: Path) -> None:
	new = not progress_csv.exists()
	with progress_csv.open('a', encoding='utf-8-sig', newline='') as f:
		w = csv.DictWriter(f, fieldnames=PROGRESS_FIELDS)
		if new:
			w.writeheader()
		c = res.customer
		w.writerow(
			{
				'Site_id': c.site_id,
				'Mã KH EVN': c.ma_kh,
				'Mã trạm biến áp': c.ma_tba,
				'Số dòng HDNSH trên web': len(res.rows),
				'Số file tải được': len(res.files),
				'Trạng thái': res.status,
				'Mã KH có trong PDF': res.verify,
				'File trùng nội dung': res.duplicates,
				'Thông tin dòng HDNSH trên web': ' | '.join(r.text for r in res.rows),
				'Thư mục': str(out_dir / c.folder_name) if res.files else '',
				'Ghi chú': '; '.join(res.notes),
				'Thời gian': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
			}
		)


def write_report(progress_csv: Path, customers: list[Customer], out_dir: Path) -> Path:
	"""Báo cáo Excel: kết quả mới nhất của từng Mã KH + các mã chưa chạy."""
	df = pd.read_csv(progress_csv, dtype=str, encoding='utf-8-sig').fillna('') if progress_csv.exists() else pd.DataFrame()
	if not df.empty:
		df = df.drop_duplicates('Mã KH EVN', keep='last')
	done = set(df['Mã KH EVN']) if not df.empty else set()
	pending = [c for c in customers if c.ma_kh not in done]
	if pending:
		rows = [
			{'Site_id': c.site_id, 'Mã KH EVN': c.ma_kh, 'Mã trạm biến áp': c.ma_tba, 'Trạng thái': 'CHƯA CHẠY'} for c in pending
		]
		df = pd.concat([df, pd.DataFrame(rows)], ignore_index=True)
	df = df.reindex(columns=PROGRESS_FIELDS).fillna('')
	summary = df['Trạng thái'].str.split(':').str[0].value_counts().rename_axis('Trạng thái').reset_index(name='Số Mã KH')
	files = pd.to_numeric(df['Số file tải được'], errors='coerce').fillna(0).sum()
	summary.loc[len(summary)] = ['Tổng số file HDNSH đã tải', int(files)]
	path = out_dir / f'BaoCao_tai_HDNSH_{datetime.now():%Y%m%d_%H%M%S}.xlsx'
	with pd.ExcelWriter(path, engine='openpyxl') as xw:
		summary.to_excel(xw, sheet_name='Tong hop', index=False)
		df.to_excel(xw, sheet_name='Chi tiet', index=False)
		for ws in xw.book.worksheets:
			for col in ws.columns:
				width = max(len(str(cell.value or '')) for cell in col[:300])
				ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 10), 60)
			ws.freeze_panes = 'A2'
	return path


def force_pdf_download(profile_dir: Path) -> None:
	"""Cấu hình profile Chromium/Edge: PDF được tải về thay vì mở trong trình xem -> bắt file ổn định hơn."""
	pref = profile_dir / 'Default' / 'Preferences'
	pref.parent.mkdir(parents=True, exist_ok=True)
	try:
		data = json.loads(pref.read_text(encoding='utf-8')) if pref.exists() else {}
	except Exception:
		data = {}
	data.setdefault('plugins', {})['always_open_pdf_externally'] = True
	data.setdefault('download', {})['prompt_for_download'] = False
	pref.write_text(json.dumps(data), encoding='utf-8')


# ----------------------------------------------------------------------------- trình duyệt


class CpcPortal:
	def __init__(self, ctx: BrowserContext, base_url: str, username: str, password: str, timeout_ms: int, debug_dir: Path):
		self.ctx = ctx
		self.base_url = base_url.rstrip('/')
		self.username = username
		self.password = password
		self.timeout_ms = timeout_ms
		self.debug_dir = debug_dir
		self.page: Page = ctx.pages[0] if ctx.pages else None  # type: ignore[assignment]
		self._captured: list[tuple[str, str, bytes]] = []
		self._capturing = False
		self._requests: list[Request] = []
		self._row_hashes: set[str] = set()
		self.search_confirmed = False
		ctx.on('response', self._on_response)
		ctx.on('request', self._on_request)

	def _on_request(self, req: Request) -> None:
		self._requests.append(req)

	async def _on_response(self, resp: Response) -> None:
		"""Bắt file trả về qua mạng khi đang mở chi tiết (popup, iframe, XHR đều đi qua context)."""
		if not self._capturing or resp.status >= 300 or resp.status == 206:
			return
		try:
			ct = (resp.headers.get('content-type') or '').lower()
			cd = (resp.headers.get('content-disposition') or '').lower()
			if 'json' in ct:
				body = await resp.body()
				for m in PDF_B64_RE.finditer(body):
					data = base64.b64decode(m.group(1).replace(b'\\r', b'').replace(b'\\n', b''))
					if data[:5] == b'%PDF-':
						self._captured.append((resp.url, 'application/pdf', data))
				return
			if not ('pdf' in ct or 'octet-stream' in ct or 'attachment' in cd or resp.url.lower().split('?')[0].endswith('.pdf')):
				return
			body = await resp.body()
			if body[:5] == b'%PDF-' or 'attachment' in cd or 'pdf' in ct:
				self._captured.append((resp.url, ct, body))
		except Exception:
			pass

	async def start(self) -> None:
		if self.page is None:
			self.page = await self.ctx.new_page()
		self.page.set_default_timeout(self.timeout_ms)
		await self.open_contract_page()

	async def open_contract_page(self) -> None:
		await self.page.goto(self.base_url + CONTRACT_PATH, wait_until='domcontentloaded')
		state = await self._wait_login_or_contract()
		if state == 'other':
			# Bị đưa về trang chủ -> bấm "Đăng nhập" nếu có
			link = self.page.get_by_role('link', name=LOGIN_TEXT).or_(self.page.get_by_role('button', name=LOGIN_TEXT))
			if await link.count():
				await link.first.click()
				state = await self._wait_login_or_contract()
		if state == 'login':
			await self.login()
			await self.page.goto(self.base_url + CONTRACT_PATH, wait_until='domcontentloaded')
			state = await self._wait_login_or_contract()
		assert state == 'contract', f'Không mở được trang Quản lý hợp đồng (URL hiện tại: {self.page.url})'

	async def _wait_login_or_contract(self) -> str:
		deadline = time.monotonic() + self.timeout_ms / 1000
		while time.monotonic() < deadline:
			if await self._is_login_page():
				return 'login'
			if CONTRACT_PATH in self.page.url and await self.page.get_by_text(re.compile(r'mã\s*khách\s*hàng', re.I)).count():
				await self._settle(5000)
				return 'contract'
			await asyncio.sleep(0.5)
		return 'other'

	async def _settle(self, ms: int = 8000) -> None:
		try:
			await self.page.wait_for_load_state('networkidle', timeout=ms)
		except PWTimeout:
			pass

	async def _is_login_page(self) -> bool:
		return await self.page.locator('input[type="password"]:visible').count() > 0

	async def login(self) -> None:
		page = self.page
		print('  Đăng nhập cskh.cpc.vn ...')
		pwd = page.locator('input[type="password"]:visible').first
		# Ô tài khoản = input text/email/tel hiển thị đứng ngay trước ô mật khẩu
		handle = await pwd.evaluate_handle(
			"""(pwd) => {
				const all = [...document.querySelectorAll('input')].filter(e => e.offsetParent !== null &&
					['text','email','tel','number',''].includes((e.getAttribute('type')||'').toLowerCase()));
				const before = all.filter(e => e.compareDocumentPosition(pwd) & Node.DOCUMENT_POSITION_FOLLOWING);
				return before.length ? before[before.length-1] : (all[0] || null);
			}"""
		)
		user = handle.as_element()
		assert user is not None, 'Không tìm thấy ô Tên đăng nhập'
		await user.fill(self.username)
		await pwd.fill(self.password)
		submit = page.get_by_role('button', name=LOGIN_TEXT)
		if await submit.count():
			await submit.first.click()
		else:
			await pwd.press('Enter')
		await self._settle()
		deadline = time.monotonic() + 300
		warned = False
		while await self._is_login_page():
			if time.monotonic() > deadline:
				raise RuntimeError('Hết 5 phút chờ đăng nhập (captcha/OTP/sai mật khẩu?)')
			if not warned:
				print(
					'  >> Chưa vào được: nếu có CAPTCHA/OTP hãy nhập trên cửa sổ trình duyệt rồi bấm Đăng nhập (chờ tối đa 5 phút).'
				)
				warned = True
			await asyncio.sleep(2)
		print('  Đăng nhập thành công.')

	# ---- chọn khách hàng & tìm kiếm
	async def search_customer(self, ma_kh: str) -> None:
		page = self.page
		if CONTRACT_PATH not in page.url or await self._is_login_page():
			await self.open_contract_page()
		box = await self._ma_kh_input()
		if box is not None and await box.is_editable():
			await box.click()
			await box.fill('')
			await box.fill(ma_kh)
		else:
			await self._select_customer_in_header(ma_kh)
			box = await self._ma_kh_input()
		prev_sig = await self._table_signature()
		btn = await self._search_button(box)
		self._requests.clear()
		if btn is not None:
			await btn.click()
		elif box is not None:
			await box.press('Enter')
		else:
			raise RuntimeError('Không tìm thấy ô Mã khách hàng / nút tìm kiếm')
		await asyncio.sleep(0.5)
		await self._settle()
		for _ in range(10):  # chờ bảng render (khác bảng cũ) tối đa ~5s
			if await self._table_signature() != prev_sig:
				break
			await asyncio.sleep(0.5)
		await asyncio.sleep(0.8)
		self.search_confirmed = any(ma_kh in (r.url + (r.post_data or '')).upper() for r in self._requests)
		shown = await self._current_ma_kh()
		if shown and shown != ma_kh:
			raise RuntimeError(f'Trang đang hiển thị Mã KH {shown} thay vì {ma_kh} (dừng để tránh lưu nhầm hồ sơ)')

	async def _ma_kh_input(self) -> Locator | None:
		page = self.page
		for cand in (
			page.get_by_label(re.compile(r'mã\s*khách\s*hàng', re.I)),
			page.get_by_placeholder(re.compile(r'mã\s*khách\s*hàng|mã\s*kh\b', re.I)),
		):
			cand = cand.locator('visible=true')
			if await cand.count():
				return cand.first
		# Nhãn "Mã khách hàng" ngoài vùng header -> input gần nhất phía sau
		ok = await page.evaluate(
			"""() => {
				document.querySelectorAll('[data-evn-makh]').forEach(e => e.removeAttribute('data-evn-makh'));
				const labels = [...document.querySelectorAll('label,span,div,p,td,th,b,strong')].filter(e =>
					e.offsetParent !== null && /^mã\\s*khách\\s*hàng\\s*:?$/i.test((e.innerText||'').trim()) && !e.closest('header'));
				for (const lb of labels) {
					let node = lb;
					for (let i = 0; i < 4 && node; i++, node = node.parentElement) {
						const inp = [...node.querySelectorAll('input')].find(x => x.offsetParent !== null &&
							(lb.compareDocumentPosition(x) & Node.DOCUMENT_POSITION_FOLLOWING));
						if (inp) { inp.setAttribute('data-evn-makh', '1'); return true; }
					}
				}
				return false;
			}"""
		)
		return page.locator('[data-evn-makh="1"]').first if ok else None

	async def _search_button(self, box: Locator | None) -> Locator | None:
		page = self.page
		named = page.get_by_role('button', name=re.compile(r'^\s*(tìm\s*kiếm|tìm|tra\s*cứu|search)\s*$', re.I))
		named = named.locator('visible=true')
		if box is None:
			return named.first if await named.count() else None
		# Nút đứng ngay sau ô Mã khách hàng (nút icon kính lúp như ảnh chụp màn hình)
		ok = await box.evaluate(
			"""(inp) => {
				document.querySelectorAll('[data-evn-search]').forEach(e => e.removeAttribute('data-evn-search'));
				let node = inp.parentElement;
				for (let i = 0; i < 5 && node; i++, node = node.parentElement) {
					const btn = [...node.querySelectorAll('button, [role=button]')].find(b => b.offsetParent !== null &&
						(inp.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING));
					if (btn) { btn.setAttribute('data-evn-search', '1'); return true; }
				}
				return false;
			}"""
		)
		if ok:
			return page.locator('[data-evn-search="1"]').first
		return named.first if await named.count() else None

	async def _select_customer_in_header(self, ma_kh: str) -> None:
		"""Trường hợp ô Mã KH bị khóa: đổi khách hàng bằng combobox Mã KH trên thanh header."""
		page = self.page
		sel = page.locator('select').filter(has=page.locator(f'option[value="{ma_kh}"]'))
		if await sel.count():
			await sel.first.select_option(ma_kh)
			await self._settle()
			return
		current = (
			page.locator('header, [class*=header], [class*=Header]').get_by_text(re.compile(r'^P[A-Z]\d{2}[A-Z]{2}\d{7}$')).first
		)
		if not await current.count():
			raise RuntimeError('Ô Mã khách hàng bị khóa và không tìm thấy combobox chọn Mã KH trên header')
		await current.click()
		await asyncio.sleep(0.5)
		typing = page.locator('input:focus')
		if await typing.count():
			await typing.fill(ma_kh)
			await asyncio.sleep(0.8)
		option = page.get_by_text(ma_kh, exact=True).locator('visible=true')
		if not await option.count():
			await page.keyboard.press('Escape')
			raise RuntimeError(f'Mã KH {ma_kh} không có trong danh sách khách hàng của tài khoản')
		await option.last.click()
		await self._settle()

	async def _current_ma_kh(self) -> str:
		box = await self._ma_kh_input()
		if box is None:
			return ''
		try:
			return (await box.input_value()).strip().upper()
		except Exception:
			return ''

	async def _table_signature(self) -> str:
		return await self.page.evaluate(
			"""() => [...document.querySelectorAll('table tbody tr, [role=row]')].filter(e => e.offsetParent !== null)
				.map(e => e.innerText.replace(/\\s+/g,' ')).join('|').slice(0, 4000)"""
		)

	async def find_hdnsh_rows(self) -> list[RowInfo]:
		"""Đánh dấu các dòng có MÃ HỒ SƠ đúng bằng HDNSH (bảng <table> hoặc lưới <div>)."""
		rows = await self.page.evaluate(
			"""(code) => {
				document.querySelectorAll('[data-evn-row]').forEach(e => e.removeAttribute('data-evn-row'));
				const isDetail = el => /xem\\s*chi\\s*ti[eế]t/i.test(el.innerText || el.getAttribute('title') || '');
				const cells = [...document.querySelectorAll('td, [role=cell], [role=gridcell], div, span, a, p')]
					.filter(e => e.offsetParent !== null && e.children.length === 0 && (e.innerText || '').trim().toUpperCase() === code);
				const out = []; const seen = new Set();
				for (const cell of cells) {
					let row = cell.closest('tr, [role=row]');
					if (!row) {
						row = cell;
						while (row && ![...row.querySelectorAll('button, a, [role=button]')].some(isDetail)) row = row.parentElement;
					}
					if (!row || seen.has(row)) continue;
					const detail = [...row.querySelectorAll('button, a, [role=button]')].filter(isDetail);
					if (detail.length !== 1) continue;  // tránh bắt nhầm container chứa nhiều dòng
					seen.add(row);
					row.setAttribute('data-evn-row', String(out.length));
					out.push(row.innerText.replace(/\\s+/g, ' ').trim());
				}
				return out;
			}""",
			MA_HO_SO,
		)
		return [RowInfo(i, t) for i, t in enumerate(rows)]

	async def data_row_count(self) -> int:
		texts = await self.page.locator('table tbody tr:visible, [role=row]:visible').all_inner_texts()
		return sum(1 for t in texts if t.strip() and not EMPTY_TEXT.search(t))

	async def has_more_pages(self) -> bool:
		"""Bảng hồ sơ có phân trang (>1 trang) -> HDNSH có thể nằm ở trang sau."""
		return await self.page.evaluate(
			"""() => [...document.querySelectorAll('[class*=paginat], [class*=Paginat], nav[aria-label*=pagination i]')]
				.filter(p => p.offsetParent !== null)
				.some(p => [...p.querySelectorAll('li, button, a')].filter(x => /^\\d+$/.test((x.innerText || '').trim())).length > 1)"""
		)

	# ---- tải file của 1 dòng HDNSH
	async def download_row(self, row: RowInfo, dest_stem: Path) -> list[Path]:
		page = self.page
		row_loc = page.locator(f'[data-evn-row="{row.index}"]')
		btn = row_loc.get_by_role('button', name=DETAIL_TEXT).or_(row_loc.get_by_role('link', name=DETAIL_TEXT))
		if not await btn.count():
			btn = row_loc.get_by_text(DETAIL_TEXT)
		btn = btn.first
		self._captured.clear()
		self._row_hashes = set()
		downloads: list[Download] = []
		popups: list[Page] = []

		def on_download(d: Download) -> None:
			downloads.append(d)

		def on_popup(p: Page) -> None:
			popups.append(p)
			p.on('download', on_download)

		page.on('download', on_download)
		self.ctx.on('page', on_popup)
		url_before = page.url
		dialogs_before = await self._dialog_count()
		saved: list[Path] = []
		self._capturing = True
		try:
			await btn.scroll_into_view_if_needed()
			await btn.click()
			saved = await self._collect(dest_stem, downloads, popups, dialogs_before)
			if not saved:
				# Chi tiết mở dạng modal / trang riêng có nút "Tải xuống"
				scopes: list[Page | Locator] = [*popups]
				if page.url != url_before:
					scopes.append(page)
				elif await self._dialog_count() > dialogs_before:
					scopes.append(page.locator(DIALOG_SEL).last)
				for scope in scopes:
					saved = await self._click_download_buttons(scope, dest_stem, downloads, popups, dialogs_before)
					if saved:
						break
			if not saved:
				saved = await self._grab_embedded(
					dest_stem, popups, include_main=page.url != url_before or await self._dialog_count() > dialogs_before
				)
		finally:
			self._capturing = False
			page.remove_listener('download', on_download)
			self.ctx.remove_listener('page', on_popup)
			for p in popups:
				if not p.is_closed():
					await p.close()
			await self._close_detail(url_before, dialogs_before)
		return saved

	async def _collect(self, dest_stem: Path, downloads: list[Download], popups: list[Page], dialogs_before: int) -> list[Path]:
		"""Chờ file về qua download hoặc response mạng; chỉ nhận file hợp lệ (không rỗng, đúng định dạng)."""
		deadline = time.monotonic() + self.timeout_ms / 1000
		opened_at: float | None = None
		while time.monotonic() < deadline:
			await asyncio.sleep(0.5)
			if downloads or self._captured:
				await asyncio.sleep(1.0)  # 1 chi tiết có thể trả nhiều file / nhiều nguồn cùng 1 file
				pending, downloads[:] = list(downloads), []
				saved = [p for d in pending if (p := self._store(await self._download_bytes(d), '', dest_stem))]
				captured, self._captured[:] = list(self._captured), []
				saved += [p for _url, ct, body in captured if (p := self._store(body, ct, dest_stem))]
				if saved:
					return saved
			if opened_at is None and (popups or await self._dialog_count() > dialogs_before):
				opened_at = time.monotonic()
			if opened_at is not None and time.monotonic() - opened_at > 8:
				return []  # chi tiết đã mở 8s mà không có file -> thử nút Tải xuống / iframe
		return []

	async def _dialog_count(self) -> int:
		return (
			await self.page.locator(DIALOG_SEL).count()
			+ await self.page.locator('iframe:visible, embed:visible, object:visible').count()
		)

	@staticmethod
	async def _download_bytes(d: Download) -> bytes:
		try:
			if await d.failure():
				return b''
			path = await d.path()
			return path.read_bytes() if path else b''
		except Exception:
			return b''

	def _store(self, data: bytes | None, content_type: str, dest_stem: Path) -> Path | None:
		"""Ghi file nếu hợp lệ và chưa trùng với file đã lưu của cùng dòng hồ sơ."""
		if not data or len(data) < 100 or data.lstrip()[:1] in (b'<', b'{', b'['):
			return None  # rỗng / trang HTML / JSON lỗi
		suffix = ext_for(content_type, data)
		if suffix == '.pdf' and data[:5] != b'%PDF-':
			return None
		h = hashlib.sha256(data).hexdigest()
		if h in self._row_hashes:
			return None
		self._row_hashes.add(h)
		n = 1
		while (target := dest_stem.with_name(f'{dest_stem.name}_{n}{suffix}')).exists():
			n += 1
		target.write_bytes(data)
		return target

	async def _click_download_buttons(
		self, scope: Page | Locator, dest_stem: Path, downloads: list[Download], popups: list[Page], dialogs_before: int
	) -> list[Path]:
		btns = scope.get_by_role('button', name=DOWNLOAD_TEXT).or_(scope.get_by_role('link', name=DOWNLOAD_TEXT))
		btns = btns.or_(scope.locator('[title*="tải" i], [aria-label*="download" i], [title*="download" i], a[download]'))
		saved: list[Path] = []
		for i in range(min(await btns.count(), 5)):
			b = btns.nth(i)
			if not await b.is_visible():
				continue
			self._captured.clear()
			downloads.clear()
			await b.click()
			saved += await self._collect(dest_stem, downloads, popups, dialogs_before + 10**6)
		return saved

	async def _grab_embedded(self, dest_stem: Path, popups: list[Page], include_main: bool) -> list[Path]:
		"""PDF hiển thị trong iframe/embed/object (kể cả blob:, data:, pdf.js ?file=) -> lấy byte trực tiếp."""
		out: list[Path] = []
		scopes = [*popups, *([self.page] if include_main else [])]
		for scope in scopes:
			srcs: list[str] = [scope.url] if scope is not self.page else []
			frames: list[Page | Frame] = [scope, *scope.frames]
			for fr in frames:
				try:
					srcs += await fr.evaluate(
						"""() => [...document.querySelectorAll('iframe[src], embed[src], object[data]')]
							.filter(e => e.offsetParent !== null || e.tagName !== 'IFRAME').map(e => e.src || e.data).filter(Boolean)"""
					)
				except Exception:
					continue
			for src in dict.fromkeys(srcs):
				data = await self._fetch_bytes(scope, src)
				if data and data[:5] == b'%PDF-' and (p := self._store(data, 'application/pdf', dest_stem)):
					out.append(p)
		return out

	async def _fetch_bytes(self, scope: Page, src: str) -> bytes | None:
		try:
			if src.startswith('data:'):
				return base64.b64decode(src.split(',', 1)[1])
			m = re.search(r'[?&]file=([^&#]+)', src)  # pdf.js viewer: .../viewer.html?file=<url>
			if m:
				src = urljoin(src, unquote(m.group(1)))
			if not src.startswith(('blob:', 'http')):
				return None
			try:  # fetch trong trang -> dùng đúng cookie/phiên của web
				b64 = await scope.evaluate(
					"""async (u) => { const r = await fetch(u, {credentials: 'include'}); if (!r.ok) return null;
						const b = new Uint8Array(await r.arrayBuffer()); let s = '';
						for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode(...b.subarray(i, i + 0x8000));
						return btoa(s); }""",
					src,
				)
				if b64:
					return base64.b64decode(b64)
			except Exception:
				pass
			if src.startswith('http'):  # bị CORS -> tải qua context (chung cookie)
				resp = await self.ctx.request.get(src)
				return await resp.body() if resp.ok else None
		except Exception:
			return None
		return None

	async def _close_detail(self, url_before: str, dialogs_before: int) -> None:
		page = self.page
		if page.url != url_before:
			try:
				await page.go_back(wait_until='domcontentloaded')
				await self._settle(4000)
			except Exception:
				pass
			return
		for _ in range(3):
			if await self._dialog_count() <= dialogs_before:
				return
			dialog = page.locator(DIALOG_SEL).last
			closers = dialog.get_by_role('button', name=CLOSE_TEXT).or_(
				dialog.locator('[aria-label="Close" i], [aria-label="Đóng" i], .ant-modal-close, .btn-close, button.close')
			)
			try:
				if await closers.count():
					await closers.first.click(timeout=3000)
				else:
					await page.keyboard.press('Escape')
			except Exception:
				await page.keyboard.press('Escape')
			await asyncio.sleep(0.6)

	async def screenshot(self, name: str) -> None:
		try:
			self.debug_dir.mkdir(parents=True, exist_ok=True)
			await self.page.screenshot(path=str(self.debug_dir / f'{safe_name(name)}.png'), full_page=True)
		except Exception:
			pass


# ----------------------------------------------------------------------------- xử lý 1 khách hàng


async def process_customer(portal: CpcPortal, c: Customer, out_dir: Path) -> CustomerResult:
	res = CustomerResult(customer=c)
	await portal.search_customer(c.ma_kh)
	if not portal.search_confirmed:
		res.notes.append('không thấy request tìm kiếm chứa Mã KH (kiểm tra lại bằng mắt)')
	res.rows = await portal.find_hdnsh_rows()
	paged = await portal.has_more_pages()
	if paged:
		res.notes.append('CẢNH BÁO: bảng hồ sơ có nhiều trang, script chỉ xử lý trang 1 -> kiểm tra các trang sau')
	if not res.rows:
		n = await portal.data_row_count()
		res.status = f'KHÔNG CÓ HDNSH: có {n} hồ sơ khác' if n else 'TRỐNG: không có hồ sơ nào (hoặc Mã KH không thuộc tài khoản)'
		if paged:
			res.status = 'CẦN KIỂM TRA: không thấy HDNSH ở trang 1, bảng có nhiều trang'
		return res
	folder = out_dir / c.folder_name
	folder.mkdir(parents=True, exist_ok=True)
	stem = folder / f'{c.folder_name}_{MA_HO_SO}'
	for old in folder.glob(f'{stem.name}_*'):  # chạy lại -> xóa file cũ để không lẫn
		old.unlink()
	for row in res.rows:
		if row.index > 0:  # sau khi đóng chi tiết, trang có thể render lại -> đánh dấu lại các dòng
			rows_now = await portal.find_hdnsh_rows()
			if len(rows_now) != len(res.rows):
				await portal.search_customer(c.ma_kh)
				rows_now = await portal.find_hdnsh_rows()
			if len(rows_now) != len(res.rows):
				res.notes.append(f'dòng {row.index + 1}: bảng thay đổi sau khi mở chi tiết')
				continue
		try:
			got = await portal.download_row(row, stem)
		except Exception as e:
			got = []
			res.notes.append(f'dòng {row.index + 1}: {type(e).__name__}: {str(e)[:150]}')
		if not got:
			res.notes.append(f'dòng {row.index + 1}: không lấy được file')
			await portal.screenshot(f'{c.folder_name}_dong{row.index + 1}')
		res.files += got

	checks = [pdf_contains(p, c.ma_kh) for p in res.files]
	if checks:
		n_ok, n_bad, n_na = checks.count(True), checks.count(False), checks.count(None)
		res.verify = (
			f'{n_ok}/{len(checks)} khớp'
			+ (f', {n_bad} KHÔNG khớp' if n_bad else '')
			+ (f', {n_na} không đọc được text' if n_na else '')
		)
	hashes: dict[str, list[str]] = {}
	for p in res.files:
		hashes.setdefault(sha256(p), []).append(p.name)
	res.duplicates = '; '.join(' = '.join(v) for v in hashes.values() if len(v) > 1)
	if not res.files:
		res.status = 'LỖI: không tải được file'
	elif len(res.files) < len(res.rows):
		res.status = f'THIẾU: {len(res.files)}/{len(res.rows)} file'
	elif False in checks:
		res.status = 'CẦN KIỂM TRA: nội dung PDF không chứa Mã KH'
	else:
		res.status = 'CẦN KIỂM TRA: bảng có nhiều trang' if paged else 'OK'
	return res


async def run(args: argparse.Namespace) -> int:
	load_dotenv(Path(__file__).with_name('.env'))
	load_dotenv()
	base_url = os.getenv('CPC_BASE_URL', 'https://cskh.cpc.vn')
	username = os.getenv('CPC_USERNAME', '')
	password = os.getenv('CPC_PASSWORD', '')
	if username and not password:
		try:
			import keyring

			password = keyring.get_password('cskh.cpc.vn', username) or ''
		except ImportError:
			pass
	if not username:
		username = input('Tên đăng nhập cskh.cpc.vn: ').strip()
	if not password:
		password = getpass.getpass(f'Mật khẩu cho {username}: ')

	only = {x.strip().upper() for x in args.only.split(',') if x.strip()} if args.only else set()
	customers = load_customers(Path(args.mapping), args.chi_nhanh, only)
	out_dir = Path(args.output).resolve()
	out_dir.mkdir(parents=True, exist_ok=True)
	progress_csv = out_dir / 'progress.csv'
	done = set() if (args.redo or only) else load_done(progress_csv)
	todo = [c for c in customers if c.ma_kh not in done]
	if args.limit:
		todo = todo[: args.limit]
	print(
		f'Mapping: {len(customers)} Mã KH ({args.chi_nhanh or "tất cả chi nhánh"}); đã xong: {len(done)}; lần này chạy: {len(todo)}'
	)
	print(f'Thư mục lưu: {out_dir}')
	if not todo:
		print('Báo cáo:', write_report(progress_csv, customers, out_dir))
		return 0

	profile_dir = out_dir / '.browser_profile'
	force_pdf_download(profile_dir)
	async with async_playwright() as pw:
		launch_kw: dict = dict(
			user_data_dir=str(profile_dir),
			headless=args.headless,
			accept_downloads=True,
			viewport={'width': 1440, 'height': 900},
			locale='vi-VN',
		)
		if args.browser_path:
			launch_kw['executable_path'] = args.browser_path
		elif args.channel:
			launch_kw['channel'] = args.channel
		try:
			ctx = await pw.chromium.launch_persistent_context(**launch_kw)
		except Exception as e:
			if 'channel' not in launch_kw:
				raise
			print(f'  Không mở được trình duyệt "{args.channel}" ({str(e)[:80]}), chuyển sang Chromium của Playwright.')
			launch_kw.pop('channel')
			ctx = await pw.chromium.launch_persistent_context(**launch_kw)
		portal = CpcPortal(ctx, base_url, username, password, args.timeout * 1000, out_dir / '_loi_chup_man_hinh')
		await portal.start()
		consecutive_err = 0
		for i, c in enumerate(todo, 1):
			t0 = time.monotonic()
			try:
				res = await process_customer(portal, c, out_dir)
			except Exception as e:
				res = CustomerResult(customer=c, status=f'LỖI: {type(e).__name__}', notes=[str(e)[:300]])
				await portal.screenshot(c.folder_name)
				try:
					await portal.open_contract_page()
				except Exception as e2:
					res.notes.append(f'mở lại trang lỗi: {str(e2)[:100]}')
			consecutive_err = consecutive_err + 1 if res.status.startswith('LỖI') else 0
			append_progress(progress_csv, res, out_dir)
			_log_customer(i, len(todo), res, time.monotonic() - t0)
			if consecutive_err >= args.max_errors:
				print(
					f'!! {consecutive_err} mã lỗi liên tiếp -> dừng để kiểm tra (ảnh chụp trong _loi_chup_man_hinh). Chạy lại sẽ tiếp tục.'
				)
				break
			await asyncio.sleep(args.delay + random.uniform(0, args.delay / 2))
		await ctx.close()
	print('Báo cáo:', write_report(progress_csv, customers, out_dir))
	return 0


def _log_customer(i: int, n: int, res: CustomerResult, secs: float) -> None:
	c = res.customer
	extra = f' | xác minh: {res.verify}' if res.verify else ''
	extra += f' | trùng: {res.duplicates}' if res.duplicates else ''
	extra += f' | {"; ".join(res.notes)}' if res.notes else ''
	print(
		f'[{i}/{n}] {c.site_id} {c.ma_kh}: {res.status} ({len(res.rows)} dòng HDNSH, {len(res.files)} file, {secs:.1f}s){extra}'
	)


def main() -> int:
	ap = argparse.ArgumentParser(description='Tải hợp đồng HDNSH từ cskh.cpc.vn theo mapping Mã KH EVN <-> Site_id')
	ap.add_argument('--mapping', required=True, help='File .xlsx/.csv có cột "Mã KH EVN", "Site_id" (tab "DS mã KH EVN")')
	ap.add_argument('--chi-nhanh', default='MobiFone Gia Lai', help='Lọc theo cột "Chi nhánh" ("" = không lọc)')
	ap.add_argument('--output', default='HD_dien_EVN_HDNSH', help='Thư mục lưu file')
	ap.add_argument('--only', default='', help='Chỉ chạy các Mã KH này (phân cách dấu phẩy)')
	ap.add_argument('--limit', type=int, default=0, help='Chỉ chạy N mã đầu tiên (chạy thử)')
	ap.add_argument('--redo', action='store_true', help='Chạy lại cả các mã đã xong')
	ap.add_argument('--delay', type=float, default=2.0, help='Giãn cách giữa 2 mã KH (giây) để không gây tải cho web EVN')
	ap.add_argument('--timeout', type=int, default=30, help='Timeout mỗi thao tác (giây)')
	ap.add_argument('--max-errors', type=int, default=5, help='Dừng khi có N mã lỗi liên tiếp')
	ap.add_argument('--headless', action='store_true', help='Chạy ẩn trình duyệt (không xử lý được captcha)')
	ap.add_argument('--channel', default='msedge' if os.name == 'nt' else '', help='Trình duyệt cài sẵn: msedge / chrome')
	ap.add_argument('--browser-path', default=os.getenv('CPC_BROWSER_PATH', ''), help='Đường dẫn file chạy Chromium/Chrome')
	args = ap.parse_args()
	if sys.platform == 'win32':
		sys.stdout.reconfigure(encoding='utf-8')  # type: ignore[attr-defined]
	return asyncio.run(run(args))


if __name__ == '__main__':
	raise SystemExit(main())
