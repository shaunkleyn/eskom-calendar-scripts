import telegram
from telegram.ext import Updater, CommandHandler, CallbackQueryHandler
import subprocess

def start(update, context):
    # Create the keyboard with the shutdown and keep alive buttons
    keyboard = [[telegram.InlineKeyboardButton("Shutdown", callback_data='shutdown'),
                telegram.InlineKeyboardButton("Keep Alive", callback_data='keep_alive')]]
    reply_markup = telegram.InlineKeyboardMarkup(keyboard)

    # Send the message with the keyboard
    update.message.reply_text('Please select an option:', reply_markup=reply_markup)

def button(update, context):
    # Get the button that was clicked
    query = update.callback_query

    # Check which button was clicked and perform the appropriate action
    if query.data == 'shutdown':
        subprocess.call('shutdown /s /t 1')
    elif query.data == 'keep_alive':
        subprocess.call('shutdown /a')
    else:
        # If the button was invalid, do nothing
        pass

# Create the Updater and pass it your bot's token.
updater = Updater("5104227651:AAHaN-7m1iWPvf-g3c5vDVIRZ0WtnR45gm0", use_context=True)

# Get the dispatcher to register handlers
dp = updater.dispatcher

# Add the start command handler
dp.add_handler(CommandHandler("start", start))

# Add the button handler
dp.add_handler(CallbackQueryHandler(button))

# Start the bot
updater.start_polling()

# Run the bot until you press Ctrl-C
updater.idle()
