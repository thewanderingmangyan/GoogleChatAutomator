import asyncio
import glob
import os
import sys
import traceback
import pandas as pd
from playwright.async_api import async_playwright

# ---------------------------------------------------------------------------
# PYINSTALLER SINGLE-FILE PATH SETUP
# ---------------------------------------------------------------------------
if getattr(sys, "frozen", False):
    base_dir = getattr(
        sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__))
    )
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(
        base_dir, "playwright", "driver", "package", ".local-browsers"
    )

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------
DOWNLOADS_DIR = os.path.expanduser("~/Downloads")
USER_DATA_DIR = os.path.expanduser("~/Desktop/google_chat_session")
SHEET_VIEW_URL = "https://docs.google.com/spreadsheets/d/1f4oi6yH__GFo6MMqIqTKoWPvKXXo9Gck6Mjo2Yo5nAE/edit?usp=sharing"
MAX_CHAT_LENGTH = 3500  # Safety threshold under Google Chat's 4096 limit


def get_latest_downloaded_csv():
    """Target and validate the latest file matching 'PENDING CRESPD AND CWET - Details'."""
    pattern = os.path.join(DOWNLOADS_DIR, "*PENDING CRESPD AND CWET*Details*")

    all_matches = glob.glob(pattern)

    completed_files = [
        f
        for f in all_matches
        if os.path.isfile(f)
        and not f.endswith((".crdownload", ".tmp", ".part"))
    ]

    if not completed_files:
        all_files = glob.glob(os.path.join(DOWNLOADS_DIR, "*Details*"))
        completed_files = [
            f
            for f in all_files
            if os.path.isfile(f)
            and not f.endswith((".crdownload", ".tmp", ".part"))
        ]

    if not completed_files:
        raise FileNotFoundError(
            f"❌ No matching file containing 'PENDING CRESPD AND CWET - Details' found in '{DOWNLOADS_DIR}'."
        )

    completed_files.sort(key=os.path.getmtime, reverse=True)

    for candidate in completed_files:
        try:
            df_head = (
                pd.read_excel(candidate, nrows=1)
                if candidate.endswith((".xlsx", ".xls"))
                else pd.read_csv(candidate, nrows=1)
            )
            headers = [str(c).strip().lower() for c in df_head.columns]

            if "fee assigned" in headers and "workordernumber" in headers:
                print(f"📄 Found latest matching Work Order file: '{candidate}'")
                return candidate
        except Exception:
            continue

    latest_file = completed_files[0]
    print(f"📄 Found latest matching Work Order file: '{latest_file}'")
    return latest_file


def parse_ageing(val):
    """Safely parses Ageing / Creation ageing values to an integer."""
    if pd.isna(val) or val is None:
        return 0

    s_val = str(val).strip().lower()

    if s_val.isdigit():
        return int(s_val)

    if "above 30" in s_val or "> 30" in s_val or ">30" in s_val:
        return 31
    if "16-30" in s_val or "16 - 30" in s_val:
        return 20
    if "8-15" in s_val or "8 - 15" in s_val:
        return 10
    if "4-7" in s_val or "4 - 7" in s_val:
        return 5

    try:
        return int(float(s_val))
    except Exception:
        return 0


def load_and_group_data(file_path):
    """Reads CSV/Excel data, filters out Completed/Cancelled WOs, targets 'Ageing' column, and groups per FEE."""
    if file_path.endswith((".xlsx", ".xls")):
        df = pd.read_excel(file_path)
    else:
        df = pd.read_csv(file_path)

    df.columns = [str(col).strip() for col in df.columns]

    column_e_name = df.columns[4] if len(df.columns) > 4 else None
    column_p_name = df.columns[15] if len(df.columns) > 15 else None
    column_q_name = df.columns[16] if len(df.columns) > 16 else None

    records = df.to_dict(orient="records")
    grouped_data = {}

    for row in records:
        fee_name = str(row.get("FEE Assigned", "")).strip()
        fee_email = str(row.get("FEE Email", "")).strip()
        wo_number = str(row.get("workordernumber", "")).strip()
        skillset = str(row.get("skillset", "Unassigned")).strip()

        status2_val = (
            str(row.get("status2", ""))
            or str(row.get("Status2", ""))
            or str(row.get("Status 2", ""))
            or str(row.get("status 2", ""))
            or (str(row.get(column_e_name, "")) if column_e_name else "")
        ).strip().lower()

        if status2_val in ["completed", "cancelled"]:
            continue

        delay_code = (
            str(row.get("delaycode", ""))
            or str(row.get("Delay Code", ""))
            or str(row.get("delay_code", ""))
        ).strip()
        if not delay_code or delay_code.lower() in ["nan", "none", ""]:
            delay_code = "None/Pending"

        ageing_val = (
            row.get("Ageing")
            or row.get("ageing")
            or (row.get(column_p_name) if column_p_name else None)
            or row.get("Creation ageing")
            or row.get("Creation Ageing")
            or (row.get(column_q_name) if column_q_name else 0)
        )
        ageing = parse_ageing(ageing_val)

        chat_sent = str(row.get("Chat Sent?", "")).strip().lower()

        if (
            not fee_name
            or fee_name.lower() in ["nan", "none"]
            or not fee_email
            or fee_email.lower() in ["nan", "none", "n/a"]
            or chat_sent == "yes"
        ):
            continue

        if fee_email not in grouped_data:
            grouped_data[fee_email] = {"name": fee_name, "work_orders": []}

        grouped_data[fee_email]["work_orders"].append(
            {
                "wo_number": wo_number,
                "skillset": skillset,
                "delay_code": delay_code,
                "ageing": ageing,
            }
        )

    return grouped_data


def format_chat_payload(fee_name, total_wos, batch, skillset_summary_str, delay_summary_str, part_info=""):
    """Constructs standardized Google Chat markdown payload for a batch of Work Orders."""
    above_30 = [w for w in batch if w["ageing"] > 30]
    days_16_to_30 = [w for w in batch if 16 <= w["ageing"] <= 30]
    days_0_to_15 = [w for w in batch if w["ageing"] <= 15]

    wo_sections = []

    if above_30:
        lines = [f"🔴 *Critical WOs for dispatch — Above 30 days ({len(above_30)} ORDERS)*"]
        for idx, wo in enumerate(above_30, start=1):
            lines.append(
                f"   {idx}. *{wo['wo_number']}* • _{wo['skillset']}_ • *{wo['delay_code']}*"
            )
        wo_sections.append("\n".join(lines))

    if days_16_to_30:
        lines = [f"🟠 *16 - 30 days ({len(days_16_to_30)} ORDERS)*"]
        for idx, wo in enumerate(days_16_to_30, start=1):
            lines.append(
                f"   {idx}. *{wo['wo_number']}* • _{wo['skillset']}_ • *{wo['delay_code']}*"
            )
        wo_sections.append("\n".join(lines))

    if days_0_to_15:
        lines = [f"🟢 *0 - 15 days ({len(days_0_to_15)} ORDERS)*"]
        for idx, wo in enumerate(days_0_to_15, start=1):
            lines.append(
                f"   {idx}. *{wo['wo_number']}* • _{wo['skillset']}_ • *{wo['delay_code']}*"
            )
        wo_sections.append("\n".join(lines))

    wo_list_body = "\n\n".join(wo_sections)

    return (
        f"🚨 *DELAYED WORK ORDER ALERT (CRESPD and CWET)*{part_info}\n"
        f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        f"Hi *{fee_name}*,\n"
        f"You have *{total_wos}* assigned Work Order(s) requiring attention:\n\n"
        f"{wo_list_body}\n"
        f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        f"📊 *ACTIVE SUMMARY*\n"
        f"  • Total Assigned: *{total_wos} Work Order(s)*\n"
        f"  • Skillsets: {skillset_summary_str}\n"
        f"  • Delay Codes: {delay_summary_str}\n\n"
        f"🔗 *Google Sheet Link:* {SHEET_VIEW_URL}\n\n"
        f"⚠️ *PLEASE ACKNOWLEDGE THIS MESSAGE*"
    )


def build_consolidated_messages(fee_name, work_orders):
    """Formats work orders into single message. Only splits into multiple parts if total character length exceeds MAX_CHAT_LENGTH (3500)."""
    total_wos = len(work_orders)

    skillsets_count = {}
    delays_count = {}

    for wo in work_orders:
        skill = wo["skillset"]
        delay = wo["delay_code"]
        skillsets_count[skill] = skillsets_count.get(skill, 0) + 1
        delays_count[delay] = delays_count.get(delay, 0) + 1

    skillset_summary_str = ", ".join(
        [f"{k}: *{v}*" for k, v in skillsets_count.items()]
    )
    delay_summary_str = ", ".join(
        [f"{k}: *{v}*" for k, v in delays_count.items()]
    )

    # 1. Attempt single-message delivery
    single_msg = format_chat_payload(
        fee_name, total_wos, work_orders, skillset_summary_str, delay_summary_str
    )

    if len(single_msg) <= MAX_CHAT_LENGTH:
        return [single_msg]

    # 2. Split only if character threshold is exceeded
    messages = []
    current_batch = []

    for wo in work_orders:
        current_batch.append(wo)
        test_msg = format_chat_payload(
            fee_name, total_wos, current_batch, skillset_summary_str, delay_summary_str, part_info=" temp"
        )
        if len(test_msg) > MAX_CHAT_LENGTH:
            overflow_wo = current_batch.pop()
            messages.append(current_batch)
            current_batch = [overflow_wo]

    if current_batch:
        messages.append(current_batch)

    total_parts = len(messages)
    final_messages = []

    for part_idx, batch in enumerate(messages, start=1):
        part_tag = f" *(Part {part_idx} of {total_parts})*"
        msg = format_chat_payload(
            fee_name, total_wos, batch, skillset_summary_str, delay_summary_str, part_info=part_tag
        )
        final_messages.append(msg)

    return final_messages


async def kill_overlays_and_popups(page):
    """Injects CSS to forcibly suppress notification pop-ups."""
    try:
        await page.add_style_tag(
            content="""
            div[role="dialog"], 
            div[aria-label*="notification"], 
            div:has-text("You are not receiving notifications"),
            .callout-bubble-class { 
                display: none !important; 
                visibility: hidden !important; 
                opacity: 0 !important; 
                pointer-events: none !important; 
            }
        """
        )
    except Exception:
        pass


async def click_new_chat(page):
    """Triggers the 'New chat' dialog."""
    selectors = [
        'span:has-text("New chat")',
        'button:has-text("New chat")',
        'div[role="button"]:has-text("New chat")',
        'button[aria-label*="New chat"]',
    ]

    for sel in selectors:
        try:
            btn = await page.wait_for_selector(sel, state="visible", timeout=3000)
            if btn:
                await btn.click()
                await asyncio.sleep(1)
                return True
        except Exception:
            continue

    await page.keyboard.press("Control+k")
    await asyncio.sleep(1)
    return True


async def select_user_suggestion_and_open_dm(page, fee_email):
    """Fills recipient email and targets user suggestion card dynamically."""
    search_selectors = [
        'input[aria-label*="Start a conversation"]',
        'input[aria-label*="Search"]',
        'input[placeholder*="person"]',
        'input[placeholder*="email"]',
        'input[role="combobox"]',
    ]

    search_input = None
    for sel in search_selectors:
        try:
            search_input = await page.wait_for_selector(
                sel, state="visible", timeout=3000
            )
            if search_input:
                break
        except Exception:
            continue

    if search_input:
        await search_input.fill(fee_email)
    else:
        await page.keyboard.type(fee_email, delay=30)

    await asyncio.sleep(1.5)

    suggestion_selectors = [
        'div[role="option"]',
        'div[data-hovercard-id]',
        f'span:has-text("{fee_email}")',
        'li[role="option"]',
    ]

    suggestion_clicked = False
    for sel in suggestion_selectors:
        try:
            option = await page.wait_for_selector(
                sel, state="visible", timeout=3000
            )
            if option:
                await option.click()
                suggestion_clicked = True
                break
        except Exception:
            continue

    if not suggestion_clicked:
        await page.keyboard.press("Enter")
        await asyncio.sleep(0.5)
        await page.keyboard.press("Enter")

    await asyncio.sleep(2)


async def send_chat_message(page, message_text):
    """Locates rich-text editor, inserts text, and dispatches message."""
    editor_selectors = [
        'div[contenteditable="true"][role="textbox"]',
        'div[aria-label*="Send a message"]',
        'div[role="textbox"][aria-multiline="true"]',
        'div[contenteditable="true"]',
    ]

    msg_box = None
    for selector in editor_selectors:
        try:
            msg_box = await page.wait_for_selector(
                selector, state="visible", timeout=6000
            )
            if msg_box:
                break
        except Exception:
            continue

    if not msg_box:
        raise Exception("Could not locate message textbox for this user.")

    await msg_box.click()
    await asyncio.sleep(0.3)

    await page.evaluate(
        """({ element, text }) => {
            element.focus();
            document.execCommand('insertText', false, text);
        }""",
        {"element": msg_box, "text": message_text},
    )

    await asyncio.sleep(0.5)
    await page.keyboard.press("Enter")
    await asyncio.sleep(1)


async def run_chat_automation():
    file_path = get_latest_downloaded_csv()
    grouped_data = load_and_group_data(file_path)

    total_recipients = len(grouped_data)
    total_wos = sum(
        len(info["work_orders"]) for info in grouped_data.values()
    )
    print(
        f"📊 Aggregated {total_wos} pending Work Orders across {total_recipients} recipient(s)."
    )

    if total_recipients == 0:
        print("ℹ️ No pending work orders to send.")
        return

    print("🌐 Launching browser context...")
    async with async_playwright() as p:
        browser = await p.chromium.launch_persistent_context(
            user_data_dir=USER_DATA_DIR,
            headless=False,
            args=[
                "--disable-notifications",
                "--disable-popup-blocking",
                "--disable-background-timer-throttling",
                "--disable-backgrounding-occluded-windows",
                "--disable-renderer-backgrounding",
            ],
            permissions=["notifications"],
        )

        page = await browser.new_page()
        print("💬 Opening Google Chat...")
        await page.goto(
            "https://chat.google.com/app/home", wait_until="domcontentloaded"
        )

        # Authentication Check
        await asyncio.sleep(3)
        if "accounts.google.com" in page.url or not (
            await page.query_selector('span:has-text("New chat")')
        ):
            print("\n" + "=" * 60)
            print(
                "🔑 ACTION REQUIRED: Please log in to your Google Workspace account inside the opened browser window."
            )
            print(
                "⏳ The automation will automatically resume once Google Chat loads..."
            )
            print("=" * 60 + "\n")

            try:
                await page.wait_for_selector(
                    'span:has-text("New chat"), button:has-text("New chat")',
                    timeout=300000,
                )
                print("✅ Login successful! Starting message dispatch...\n")
            except Exception:
                raise TimeoutError(
                    "❌ Login timeout exceeded (5 minutes). Please run the application again."
                )

        await kill_overlays_and_popups(page)

        # Results Tracker
        dispatch_report = []

        for fee_email, data in grouped_data.items():
            if page.is_closed():
                print("⚠️ Browser tab closed. Exiting.")
                break

            fee_name = data["name"]
            work_orders = data["work_orders"]
            message_batches = build_consolidated_messages(fee_name, work_orders)

            print(
                f"🚀 Processing DM to {fee_email} ({len(work_orders)} WOs in {len(message_batches)} batch[es])..."
            )

            user_success = True
            batches_sent = 0

            try:
                await kill_overlays_and_popups(page)
                await click_new_chat(page)
                await select_user_suggestion_and_open_dm(page, fee_email)

                for msg in message_batches:
                    await send_chat_message(page, msg)
                    batches_sent += 1
                    await asyncio.sleep(1)

                print(f"✅ Delivered all DM batches to {fee_email}")

            except Exception as e:
                user_success = False
                print(f"❌ Unsuccessful dispatch to {fee_email}: {str(e)}")

            dispatch_report.append(
                {
                    "recipient": fee_email,
                    "name": fee_name,
                    "status": "Delivered" if user_success else "Unsuccessful",
                    "batches": batches_sent,
                    "total_wos": len(work_orders),
                }
            )

            await asyncio.sleep(1.5)

        # Execution Summary Table
        print("\n" + "=" * 60)
        print("📊 EXECUTION DISPATCH SUMMARY REPORT")
        print("=" * 60)
        print(f"{'Recipient Email':<35} | {'Status':<15} | {'WOs':<5} | {'DMs'}")
        print("-" * 60)

        successful_count = 0
        for item in dispatch_report:
            status_str = (
                "✅ Delivered"
                if item["status"] == "Delivered"
                else "❌ Unsuccessful"
            )
            if item["status"] == "Delivered":
                successful_count += 1
            print(
                f"{item['recipient']:<35} | {status_str:<15} | {item['total_wos']:<5} | {item['batches']}"
            )

        print("=" * 60)
        print(
            f"🎉 Completed: {successful_count} of {len(dispatch_report)} recipient(s) successfully processed.\n"
        )
        await browser.close()


if __name__ == "__main__":
    try:
        asyncio.run(run_chat_automation())
    except Exception as e:
        print(f"\n❌ CRITICAL SCRIPT ERROR: {str(e)}")
        traceback.print_exc()
    finally:
        print("\n" + "=" * 50)
        input("Press ENTER to exit...")
