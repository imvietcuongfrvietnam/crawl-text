import os
import re
import time
import shutil
import json
import anthropic
import pandas as pd
import fitz  # PyMuPDF
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    InvalidSessionIdException, WebDriverException,
    TimeoutException, NoSuchElementException,
)
from webdriver_manager.chrome import ChromeDriverManager

# ============================================================
# CẤU HÌNH
# ============================================================
INPUT_FILE  = 'raw.xlsx'
INPUT_SHEET = '1. Raw data'
BASE_DATA   = os.path.join(os.getcwd(), 'data')
TEMP_DL     = os.path.join(os.getcwd(), '_temp_dl')

for _d in (BASE_DATA, TEMP_DL):
    os.makedirs(_d, exist_ok=True)

# ============================================================
# CHROME
# ============================================================
def make_chrome_options():
    opt = Options()
    # opt.add_argument('--headless')  # bỏ comment để chạy ngầm
    opt.add_experimental_option("prefs", {
        "download.default_directory": TEMP_DL,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "safebrowsing.enabled": False,
        "safebrowsing.disable_download_protection": True,
    })
    opt.add_argument("--disable-features=DownloadBubble,DownloadBubbleV2")
    opt.add_argument("--no-first-run")
    opt.add_argument("--no-default-browser-check")
    return opt

def make_driver():
    svc = Service(ChromeDriverManager().install())
    d = webdriver.Chrome(service=svc, options=make_chrome_options())
    d.implicitly_wait(2)
    return d

# ============================================================
# TOOL DEFINITIONS (schema chỉ dùng để gọi Claude API)
# ============================================================
TOOLS = [
    {
        "name": "navigate_to_url",
        "description": (
            "Navigate the browser to a URL and wait for the page to load. "
            "Returns page title and current URL."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "url": {"type": "string", "description": "Full URL to navigate to"},
                "wait_seconds": {
                    "type": "number",
                    "description": "Extra seconds to wait after page loads (default 2)"
                }
            },
            "required": ["url"]
        }
    },
    {
        "name": "get_page_content",
        "description": (
            "Return a structured, readable snapshot of the current page: "
            "tables (with cell text), all clickable elements (buttons/links/tabs), "
            "and headings. Use this to understand page layout before clicking. "
            "Pass an optional CSS selector to limit to a part of the page."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "css_scope": {
                    "type": "string",
                    "description": "Optional CSS selector to limit scope (e.g. '#content', '.tab-panel')"
                }
            }
        }
    },
    {
        "name": "click_element",
        "description": (
            "Click an element identified by XPath or CSS selector. "
            "After clicking, the page may change or a new tab may open."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "xpath": {
                    "type": "string",
                    "description": "XPath expression to locate the element"
                },
                "css_selector": {
                    "type": "string",
                    "description": "CSS selector to locate the element (alternative to xpath)"
                },
                "timeout": {
                    "type": "integer",
                    "description": "Max seconds to wait for element to be clickable (default 10)"
                }
            }
        }
    },
    {
        "name": "check_tabs",
        "description": (
            "Check how many browser tabs are currently open. "
            "If switch_to_new is true and a new tab exists, switch to it and return its URL/title."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "switch_to_new": {
                    "type": "boolean",
                    "description": "If true, switch to the new tab (non-main tab)"
                }
            }
        }
    },
    {
        "name": "close_current_tab",
        "description": "Close the current tab and switch back to the main (first) tab.",
        "input_schema": {"type": "object", "properties": {}}
    },
    {
        "name": "clear_temp_dir",
        "description": "Delete all files in the temp download folder before starting a new download.",
        "input_schema": {"type": "object", "properties": {}}
    },
    {
        "name": "check_downloads",
        "description": (
            "List files in the temp download folder. "
            "Returns completed files (no .crdownload/.tmp suffix) and in-progress files."
        ),
        "input_schema": {"type": "object", "properties": {}}
    },
    {
        "name": "wait_for_download",
        "description": (
            "Repeatedly check the temp folder until a completed file appears, "
            "or until timeout. Returns filename if found, or 'timeout' message."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "timeout": {
                    "type": "integer",
                    "description": "Max seconds to wait (default 15)"
                }
            }
        }
    },
    {
        "name": "move_downloaded_file",
        "description": (
            "Move a file from the temp download folder to the current package folder "
            "with a given base name (extension is preserved)."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "filename": {
                    "type": "string",
                    "description": "Exact filename that currently exists in temp folder"
                },
                "dest_name": {
                    "type": "string",
                    "description": (
                        "Base name for destination file without extension, "
                        "e.g. 'bieu_mau_moi_thau' or 'yeu_cau_ky_thuat'"
                    )
                }
            },
            "required": ["filename", "dest_name"]
        }
    },
    {
        "name": "extract_pdf_section",
        "description": (
            "Extract the 'PHẠM VI CUNG CẤP' (danh mục hàng hoá) section from a PDF. "
            "Returns the section text, or the full PDF text if the section is not found."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Absolute path to the PDF file"
                }
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "sleep",
        "description": "Pause execution for the given number of seconds (max 15).",
        "input_schema": {
            "type": "object",
            "properties": {
                "seconds": {"type": "number", "description": "Seconds to wait"}
            },
            "required": ["seconds"]
        }
    },
    {
        "name": "finish_package",
        "description": (
            "Call this when done processing the current package. "
            "Records the extracted PHẠM VI CUNG CẤP / danh mục hàng hoá content "
            "and terminates the agent loop for this package."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "content": {
                    "type": "string",
                    "description": "Extracted danh mục hàng hoá text (or error message)"
                }
            },
            "required": ["content"]
        }
    }
]

# ============================================================
# PAGE SNAPSHOT — structured content for AI
# ============================================================
_PAGE_SCRIPT = """
(function(scopeSelector) {
    var root = scopeSelector
        ? (document.querySelector(scopeSelector) || document.body)
        : document.body;

    var lines = [];
    lines.push('URL: ' + window.location.href);
    lines.push('TITLE: ' + document.title);
    lines.push('');

    // --- Tables ---
    var tables = root.querySelectorAll('table');
    if (tables.length) {
        lines.push('=== TABLES ===');
        tables.forEach(function(tbl, ti) {
            lines.push('-- Table ' + (ti + 1) + ' --');
            tbl.querySelectorAll('tr').forEach(function(tr) {
                var cells = Array.from(tr.querySelectorAll('th,td')).map(function(td) {
                    var txt = (td.textContent || '').trim().replace(/\\s+/g, ' ');
                    // collect chip/badge/file elements inside the cell
                    var chips = Array.from(
                        td.querySelectorAll('a, button, [class*="chip"], [class*="badge"], [class*="file"], span')
                    ).map(function(el) {
                        var t = (el.textContent || '').trim().replace(/\\s+/g, ' ');
                        var href = el.getAttribute('href') || '';
                        return t ? (href ? '[LINK:' + t + '|' + href + ']' : '[EL:' + t + ']') : '';
                    }).filter(Boolean);
                    // only show chips if they add info not in the cell text
                    var unique = chips.filter(function(c) {
                        return txt.indexOf(c.replace(/\\[.*?:/,'').replace(/\\|.*\\]/,']').replace(/\\]/,'')) === -1;
                    });
                    return unique.length ? txt + ' ' + unique.join(' ') : txt;
                });
                if (cells.some(function(c) { return c.trim(); })) {
                    lines.push('| ' + cells.join(' | ') + ' |');
                }
            });
        });
        lines.push('');
    }

    // --- Tabs / Nav ---
    var tabs = root.querySelectorAll('[role="tab"], .nav-item, .tab-item, .tab, li.active, li > a');
    if (tabs.length) {
        lines.push('=== TABS / NAV ===');
        var seen = {};
        Array.from(tabs).forEach(function(el) {
            var t = (el.textContent || '').trim().replace(/\\s+/g, ' ');
            if (t && !seen[t]) { seen[t] = 1; lines.push('  [TAB] ' + t); }
        });
        lines.push('');
    }

    // --- Buttons ---
    var btns = root.querySelectorAll('button, [role="button"]');
    if (btns.length) {
        lines.push('=== BUTTONS ===');
        var seen2 = {};
        Array.from(btns).forEach(function(el) {
            var t = (el.textContent || '').trim().replace(/\\s+/g, ' ');
            if (t && !seen2[t]) { seen2[t] = 1; lines.push('  [BTN] ' + t); }
        });
        lines.push('');
    }

    // --- Links ---
    var links = root.querySelectorAll('a[href]');
    if (links.length) {
        lines.push('=== LINKS ===');
        var seen3 = {};
        Array.from(links).forEach(function(el) {
            var t = (el.textContent || '').trim().replace(/\\s+/g, ' ');
            var href = el.getAttribute('href') || '';
            if (t && !seen3[t] && t.length < 300) {
                seen3[t] = 1;
                lines.push('  [A] ' + t + (href ? ' | ' + href : ''));
            }
        });
        lines.push('');
    }

    // --- Headings ---
    var hds = root.querySelectorAll('h1,h2,h3,h4,h5');
    if (hds.length) {
        lines.push('=== HEADINGS ===');
        hds.forEach(function(h) {
            lines.push('  [' + h.tagName + '] ' + (h.textContent || '').trim().replace(/\\s+/g, ' '));
        });
    }

    return lines.join('\\n');
})(arguments[0]);
"""

def get_page_snapshot(driver, css_scope=None):
    try:
        raw = driver.execute_script(_PAGE_SCRIPT, css_scope or None)
        return str(raw)[:40000]
    except Exception as e:
        return f"Error getting page content: {e}"

# ============================================================
# TOOL EXECUTOR
# ============================================================
def execute_tool(name, inp, state):
    driver = state['driver']
    main_win = state.get('main_window')
    pkg_folder = state['pkg_folder']

    try:
        # ---- navigate_to_url ----
        if name == "navigate_to_url":
            url = inp['url']
            wait_sec = float(inp.get('wait_seconds', 2))
            driver.get(url)
            state['main_window'] = driver.current_window_handle
            time.sleep(min(wait_sec, 10))
            return f"OK. Title: {driver.title}\nURL: {driver.current_url}"

        # ---- get_page_content ----
        elif name == "get_page_content":
            return get_page_snapshot(driver, inp.get('css_scope'))

        # ---- click_element ----
        elif name == "click_element":
            timeout = int(inp.get('timeout', 10))
            wait = WebDriverWait(driver, timeout)
            xpath = inp.get('xpath', '').strip()
            css = inp.get('css_selector', '').strip()
            if xpath:
                el = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            elif css:
                el = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, css)))
            else:
                return "Error: provide xpath or css_selector"
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'}); arguments[0].click();", el
            )
            time.sleep(1)
            return f"Clicked. Current URL: {driver.current_url}"

        # ---- check_tabs ----
        elif name == "check_tabs":
            handles = driver.window_handles
            mw = state.get('main_window', handles[0] if handles else None)
            others = [h for h in handles if h != mw]
            info = f"Total tabs: {len(handles)}\nNew tabs: {len(others)}"
            if inp.get('switch_to_new') and others:
                driver.switch_to.window(others[0])
                info += f"\nSwitched. URL: {driver.current_url}\nTitle: {driver.title}"
            return info

        # ---- close_current_tab ----
        elif name == "close_current_tab":
            mw = state.get('main_window')
            if mw and driver.current_window_handle != mw:
                driver.close()
                driver.switch_to.window(mw)
                return f"Tab closed. Back on main tab. URL: {driver.current_url}"
            return "Already on main tab."

        # ---- clear_temp_dir ----
        elif name == "clear_temp_dir":
            count = 0
            for f in os.listdir(TEMP_DL):
                try:
                    os.remove(os.path.join(TEMP_DL, f))
                    count += 1
                except Exception:
                    pass
            return f"Cleared {count} file(s) from temp dir."

        # ---- check_downloads ----
        elif name == "check_downloads":
            all_files = os.listdir(TEMP_DL)
            done  = [f for f in all_files if not f.endswith(('.crdownload', '.tmp'))]
            prog  = [f for f in all_files if f.endswith(('.crdownload', '.tmp'))]
            return f"Completed: {done}\nIn progress: {prog}"

        # ---- wait_for_download ----
        elif name == "wait_for_download":
            timeout = int(inp.get('timeout', 15))
            for _ in range(timeout):
                time.sleep(1)
                done = [f for f in os.listdir(TEMP_DL)
                        if not f.endswith(('.crdownload', '.tmp'))]
                if done:
                    return f"Downloaded: {done[0]}"
            return "Timeout: no completed file found."

        # ---- move_downloaded_file ----
        elif name == "move_downloaded_file":
            fn = inp['filename']
            dest_name = inp['dest_name']
            src = os.path.join(TEMP_DL, fn)
            if not os.path.exists(src):
                return f"Error: '{fn}' not found in temp dir. Files: {os.listdir(TEMP_DL)}"
            ext = os.path.splitext(fn)[1]
            dst = os.path.join(pkg_folder, f"{dest_name}{ext}")
            if os.path.exists(dst):
                os.remove(dst)
            shutil.move(src, dst)
            return f"Moved to: {dst}"

        # ---- extract_pdf_section ----
        elif name == "extract_pdf_section":
            fp = inp['file_path']
            if not os.path.exists(fp):
                return f"Error: file not found: {fp}"
            full = ""
            with fitz.open(fp) as doc:
                for page in doc:
                    full += page.get_text()
            pat = re.compile(r'PH[ẠA]M VI CUNG C[ẤA]P', re.IGNORECASE)
            m = pat.search(full)
            if not m:
                return full.strip()[:30000]
            section = full[m.start():]
            end_pat = re.compile(
                r'\nM[ẫẪ]u s[ốỐ]\s+\d|\nCh[ươưƯ][oO]ng\s+[IVXLC\d]|'
                r'\nPh[ầẦ]n\s+\d|\nBi[ểỂ]u m[ẫẪ]u|\nYêu c[ầẦ]u',
                re.IGNORECASE
            )
            em = end_pat.search(section, 50)
            if em:
                section = section[:em.start()]
            return section.strip()[:30000]

        # ---- sleep ----
        elif name == "sleep":
            secs = min(float(inp.get('seconds', 2)), 15)
            time.sleep(secs)
            return f"Slept {secs}s"

        # ---- finish_package ----
        elif name == "finish_package":
            content = inp.get('content', '')
            state['result']['content'] = content
            state['result']['done'] = True
            return f"Package finished. Content length: {len(content)}"

        else:
            return f"Unknown tool: {name}"

    except TimeoutException:
        return "Timeout: element not found or not clickable."
    except Exception as e:
        return f"Tool error ({name}): {str(e).split(chr(10))[0]}"

# ============================================================
# SYSTEM PROMPT
# ============================================================
SYSTEM_PROMPT = """You are an autonomous web scraping agent for Vietnamese government procurement data at muasamcong.mpi.gov.vn.

## Your task for each package:

1. **Navigate** to the package URL.
2. **Click the "Hồ sơ mời thầu" tab** — use get_page_content to find it, then click_element.
3. **Download Biểu mẫu mời thầu (row 1.5)**:
   - Find the row with STT "1.5" and text "Biểu mẫu mời thầu"
   - Click the file chip/link in that row
   - A new viewer tab typically opens — use check_tabs + switch_to_new=true
   - In the viewer, find and click the "Tải về" button
   - Call wait_for_download, then move_downloaded_file with dest_name="bieu_mau_moi_thau"
   - If the file is a PDF, call extract_pdf_section with the full path returned by move_downloaded_file
4. **Return to main tab** if needed (close_current_tab).
5. **Download Yêu cầu kỹ thuật files (Chương V)**:
   - Find the "Chương V" row in the table (may be labeled "Chương V", or have STT like "2", "V", etc.)
   - Check if it has sub-rows (child rows with STT like "2.1", "2.1.1", etc.)
   - **If NO sub-rows**: download all file chips from the Chương V row itself
   - **If HAS sub-rows**: only download files from rows containing "Yêu cầu kỹ thuật" or "Chỉ dẫn kỹ thuật"
   - Save each file as "yeu_cau_ky_thuat" (or "yeu_cau_ky_thuat_1", "_2" for multiple files)
   - For each file chip: clear_temp_dir, click_element, then poll check_tabs + check_downloads every ~1s
     - If new tab: switch + find "Tải về" button + wait_for_download
     - If direct download: wait_for_download
6. **Call finish_package** with the extracted PHẠM VI CUNG CẤP content (or an error message).

## Important notes:
- Always call get_page_content before navigating — understand the structure first.
- File chips are clickable <a> or <span> elements inside table cells, often with a file extension in their text (.pdf, .docx, .xlsx, etc.)
- After clicking a chip: immediately poll check_tabs (switch_to_new=true) — if a tab opens, handle the viewer. Otherwise use wait_for_download.
- The page is in Vietnamese. Read text carefully to identify the correct rows and elements.
- If a step fails, continue with remaining steps. Always call finish_package at the end.
- When calling click_element, prefer XPath with normalize-space() for text matching, e.g.:
    //a[normalize-space(text())='Tải về']
    //*[normalize-space(.)='Hồ sơ mời thầu']
    //tr[td[normalize-space(text())='1.5']]//a
"""

# ============================================================
# AGENT LOOP (per package)
# ============================================================
def run_agent(client, driver, ma_tbmt, url, pkg_folder):
    state = {
        'driver': driver,
        'main_window': driver.current_window_handle if driver.window_handles else None,
        'pkg_folder': pkg_folder,
        'result': {'content': '', 'done': False},
    }

    messages = [{
        "role": "user",
        "content": (
            f"Process procurement package:\n"
            f"- Mã TBMT: {ma_tbmt}\n"
            f"- URL: {url}\n"
            f"- Package folder: {pkg_folder}\n\n"
            f"Follow the system instructions. Call finish_package when done."
        )
    }]

    MAX_ITER = 50
    for iteration in range(MAX_ITER):
        response = client.messages.create(
            model="claude-opus-4-7",
            max_tokens=8192,
            thinking={"type": "adaptive"},
            system=SYSTEM_PROMPT,
            tools=TOOLS,
            messages=messages,
        )

        # Append assistant turn (including thinking blocks)
        messages.append({"role": "assistant", "content": response.content})

        if response.stop_reason == "end_turn":
            break

        if response.stop_reason != "tool_use":
            print(f"      [Agent] stop_reason={response.stop_reason}")
            break

        if state['result'].get('done'):
            break

        # Execute all tool calls in the response
        tool_results = []
        for block in response.content:
            if block.type != "tool_use":
                continue
            tool_name = block.name
            tool_input = block.input
            short_input = {k: str(v)[:120] for k, v in tool_input.items()}
            print(f"      [Tool] {tool_name}({json.dumps(short_input, ensure_ascii=False)})")
            result_text = execute_tool(tool_name, tool_input, state)
            # Truncate large results (page content can be large but we still need it)
            max_len = 40000 if tool_name == "get_page_content" else 8000
            if len(result_text) > max_len:
                result_text = result_text[:max_len] + "\n...[truncated]"
            tool_results.append({
                "type": "tool_result",
                "tool_use_id": block.id,
                "content": result_text,
            })

        if tool_results:
            messages.append({"role": "user", "content": tool_results})

        if state['result'].get('done'):
            break

    return state['result'].get('content', 'Agent did not call finish_package')

# ============================================================
# MAIN
# ============================================================
def main():
    print(f"Đang đọc file: {INPUT_FILE}...")
    try:
        df = pd.read_excel(INPUT_FILE, sheet_name=INPUT_SHEET)
    except Exception as e:
        print(f"Lỗi đọc file: {e}")
        return

    df = df.drop_duplicates(subset=['Mã TBMT'], keep='first').reset_index(drop=True)
    print(f"Số bản ghi cần xử lý: {len(df)}")

    client = anthropic.Anthropic()  # reads ANTHROPIC_API_KEY from env
    driver = make_driver()
    content_list = []

    for index, row in df.iterrows():
        url     = row['Link chi tiết']
        ma_tbmt = str(row['Mã TBMT']).strip()
        print(f"\n[{index + 1}/{len(df)}] {ma_tbmt}")

        if pd.isna(url) or str(url).strip() == "":
            print("   -> Link trống, bỏ qua.")
            content_list.append("Không có link")
            continue

        pkg_folder = os.path.join(BASE_DATA, ma_tbmt)
        os.makedirs(pkg_folder, exist_ok=True)

        # Restart driver if session died
        try:
            driver.current_url
        except (InvalidSessionIdException, WebDriverException):
            print("   !! Session died — khởi động lại trình duyệt...")
            try:
                driver.quit()
            except Exception:
                pass
            driver = make_driver()

        try:
            content = run_agent(client, driver, ma_tbmt, url, pkg_folder)
            content_list.append(content)
            print(f"   -> Done. Content: {len(content)} chars")
        except Exception as e:
            err = str(e).split('\n')[0]
            print(f"   -> Lỗi agent: {err}")
            content_list.append(f"Lỗi: {err}")

    try:
        driver.quit()
    except Exception:
        pass

    df['content'] = [str(c)[:30000] for c in content_list]
    output_file = 'ket_qua_cuoi_cung.xlsx'
    df.to_excel(output_file, index=False)
    print(f"\n======================================")
    print(f"Xong! Kết quả: {output_file}")
    print(f"File gốc: data/{{ma_tbmt}}/bieu_mau_moi_thau.*")
    print(f"         data/{{ma_tbmt}}/yeu_cau_ky_thuat.*")
    print(f"======================================")


if __name__ == "__main__":
    main()
