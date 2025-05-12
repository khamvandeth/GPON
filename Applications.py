import requests
import openpyxl
import logging
import re
from io import BytesIO
from telegram import Update, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
    ConversationHandler
)

# Configure logging
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# ===== Configuration =====
BOT_TOKEN = '8002030549:AAGgPhsN4CTbfmzXwk4-PxUkXUxt6D2krbQ'  # Replace with your actual token
GITHUB_EXCEL_URL = "https://github.com/khamvandeth/GPON/raw/main/GPON.xlsx"
BCCS_API_URL = "http://36.37.242.67:8068/BCCSGatewayWS/BCCSGatewayWS?wsdl"

# Columns to exclude from GPON display
EXCLUDED_COLUMNS = {'No', 'BTS Name'}

# Conversation states
CHOOSING, Search_Site, Change_Device = range(3)

# ===== Shared Data =====
excel_data = None

# ===== Helper Functions =====
def load_excel_data():
    global excel_data
    try:
        response = requests.get(GITHUB_EXCEL_URL, headers={
            'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        })
        response.raise_for_status()
        
        workbook = openpyxl.load_workbook(filename=BytesIO(response.content))
        sheet = workbook.active
        
        data = []
        headers = [cell.value for cell in sheet[1]]
        
        for row in sheet.iter_rows(min_row=2, values_only=True):
            data.append(dict(zip(headers, row)))
            
        excel_data = data
        return data
    except Exception as error:
        logger.error(f'Error loading Excel file: {str(error)}')
        return None

def search_data(data, term):
    if not data:
        return []
    term = term.lower()
    results = []
    for row in data:
        if any(term in str(value).lower() for value in row.values()):
            results.append(row)
    return results

def format_result(result):
    formatted = []
    for key, value in result.items():
        if key not in EXCLUDED_COLUMNS:
            formatted.append(f"<b>{key}:</b> <code>{value}</code>")
    return '\n'.join(formatted)

# ===== Command Handlers =====
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Send welcome message and main menu."""
    reply_keyboard = [
        [KeyboardButton("Search Site"), KeyboardButton("Change Device")],
        [KeyboardButton("Help")]
    ]
    
    welcome_text = """
✨ <b>Welcome to Combined Telecom Bot!</b> ✨

Please choose an operation:
- <b>Search Site</b>: Search Site
- <b>Change Device</b>: Change device for account

You can type /Back at any time to return to this menu.
"""
    await update.message.reply_text(
        welcome_text,
        parse_mode='HTML',
        reply_markup=ReplyKeyboardMarkup(
            reply_keyboard, 
            resize_keyboard=True,
            one_time_keyboard=True
        )
    )
    return CHOOSING

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Send help message."""
    help_text = """
ℹ️ <b>Combined Telecom Bot Help</b> ℹ️

<b>Search Site</b>:
Send any search term to find matching records in GPON database./Back

<b>Change Device</b>:
Send your request in this format:
<code>account:YOUR_ACCOUNT
device:YOUR_DEVICE_CODE</code>

Example:
<code>account:98xxxxxxxxx
device:PNP111_A_G_C610</code>
"""
    await update.message.reply_text(help_text, parse_mode='HTML')
    return CHOOSING

async def Back(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Return to main menu."""
    await start(update, context)
    return CHOOSING

# ===== Search Site Handlers =====
async def search_site(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Start Search Site mode."""
    await update.message.reply_text(
        "🔍 <b>Search Site Mode</b>\n\n"
        "Enter your search term (SITE, IP, etc.)\n"
        "Type /Back to return to main menu.",
        parse_mode='HTML'
    )
    return Search_Site

async def handle_gpon_search(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle Search Site queries."""
    global excel_data
    try:
        if not excel_data:
            await update.message.reply_text('⏳ Loading data...')
            excel_data = load_excel_data()
            if not excel_data:
                await update.message.reply_text('❌ Failed to load data. Please try again later.')
                return await Back(update, context)
        
        search_term = update.message.text
        results = search_data(excel_data, search_term)
        
        if results:
            header = f"🔍 Found <b>{len(results)}</b> matches for '<code>{search_term}</code>':\n\n"
            formatted_results = '\n\n'.join([format_result(result) for result in results[:5]])
            
            if len(results) > 5:
                more_text = f"\n\n...and <b>{len(results) - 5}</b> more results not shown"
            else:
                more_text = ""
                
            footer = "\n\nℹ️ Tip: Try a more specific search term for better results./Back"
            
            full_message = header + formatted_results + more_text + footer
            await update.message.reply_text(full_message, parse_mode='HTML')
        else:
            no_results_text = f"""
❌ No matches found for '<code>{search_term}</code>'

Suggestions:
- Check for typos
- Try different keywords
- Be less specific
"""
            await update.message.reply_text(no_results_text, parse_mode='HTML')
            
    except Exception as error:
        logger.error(f'Error handling Search Site: {error}')
        error_text = "⚠️ <b>An error occurred</b>\nWe couldn't process your request. Please try again."
        await update.message.reply_text(error_text, parse_mode='HTML')
    
    return Search_Site

# ===== Change Device Handlers =====
async def change_device(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Start Change Device mode."""
    await update.message.reply_text(
        "🔄 <b>Change Device Mode</b>\n\n"
        "Send your request in this format:\n"
        "<code>account:YOUR_ACCOUNT\ndevice:YOUR_DEVICE_CODE</code>\n\n"
        "Example:\n"
        "<code>account:98xxxxxxxxx\ndevice:PNP111_A_G_C610</code>\n\n"
        "Type /Back to return to main menu.",
        parse_mode='HTML'
    )
    return Change_Device

async def handle_bccs_operation(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle Change Device requests."""
    message_text = update.message.text.strip()
    
    try:
        # Parse account and deviceCode from message
        account_match = re.search(r'account:(\S+)', message_text)
        device_match = re.search(r'device:(\S+)', message_text)
        
        if not account_match or not device_match:
            await update.message.reply_text(
                "⚠️ <b>Invalid Format</b> ⚠️\n\n"
                "Please use:\n"
                "<code>account:YOUR_ACCOUNT\ndevice:YOUR_DEVICE_CODE</code>\n\n"
                "<b>Example:</b>\n"
                "<code>account:98xxxxxxxxx\ndevice:PNP111_A_G_C610</code>",
                parse_mode='HTML'
            )
            return Change_Device
            
        account = account_match.group(1)
        device_code = device_match.group(1)
        
        # Prepare the SOAP payload
        payload = f"""<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:web="http://webservice.bccsgw.viettel.com/">
           <soapenv:Header/>
           <soapenv:Body>
              <web:gwOperation>
                 <Input>
                    <username>addba2e908c412ca</username>
                    <password>523cc9765677493c4a2fe6ef8b80d222</password>
                    <wscode>ChangeDeviceAccFtth</wscode>                
                    <param name="token" value="c1u1o1n1g143045ef95bb959ab2448f9072c086c90d01a4"/>       
                    <param name="locale" value="en_US"/>
                    <param name="account" value="{account}"/>
                    <param name="deviceCode" value="{device_code}"/> 
                </Input>
              </web:gwOperation>
           </soapenv:Body>
        </soapenv:Envelope>"""
        
        # Send to API
        headers = {
            'Content-Type': 'text/xml',
            'SOAPAction': ''
        }
        response = requests.post(BCCS_API_URL, headers=headers, data=payload)
        
        logger.info(f"API response: {response.text}")
        
        # Check response status
        if "success" in response.text.lower():
            status = "✅ Success"
        elif "can not find task for" in response.text.lower():
            status = "⚠️ Can not find task for account"
        elif "not find device" in response.text.lower():
            status = "⚠️ Can not find device"
        else:
            status = "❌ Unknown response"
        
        # Format the response
        formatted_response = (
            f"<b>Request Status:</b> {status}\n\n"
            f"🔹 <b>Account:</b> <code>{account}</code>\n"
            f"🔹 <b>Device Code:</b> <code>{device_code}</code>\n\n"
            "<i>The request has been processed by the API./Back</i>"
        )
        
        await update.message.reply_text(formatted_response, parse_mode='HTML')
        
    except Exception as e:
        logger.error(f"Error processing Change Device: {e}")
        await update.message.reply_text(
            f"❌ <b>Error Processing Request</b> ❌\n\n"
            f"<code>{str(e)}</code>",
            parse_mode='HTML'
        )
    
    return Change_Device

# ===== Main Function =====
def main():
    """Start the bot."""
    application = Application.builder().token(BOT_TOKEN).build()

    # Set up conversation handler with the states
    conv_handler = ConversationHandler(
        entry_points=[CommandHandler('start', start)],
        states={
            CHOOSING: [
                MessageHandler(filters.Regex('^Search Site$'), search_site),
                MessageHandler(filters.Regex('^Change Device$'), change_device),
                MessageHandler(filters.Regex('^Help$'), help_command),
            ],
            Search_Site: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, handle_gpon_search),
                CommandHandler('Back', Back),
            ],
            Change_Device: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, handle_bccs_operation),
                CommandHandler('Back', Back),
            ],
        },
        fallbacks=[CommandHandler('Back', Back)],
    )

    application.add_handler(conv_handler)
    application.add_handler(CommandHandler('help', help_command))

    logger.info("Bot is running and waiting for messages...")
    application.run_polling()

if __name__ == "__main__":
    main()
