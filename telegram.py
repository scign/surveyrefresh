import os
from telethon import TelegramClient
from sharepoint import get_file_details

async def send_message(message):
  async with TelegramClient(
      session='survey_notification',
      api_id=os.environ.get('TG_API_ID'),
      api_hash=os.environ.get('TG_API_HASH')) as tc:
    bot = await tc.start(bot_token=os.environ.get('TG_BOT_TOKEN'))
    dst = await tc.get_entity(int(os.environ.get('TG_GROUP_ID')))
    await bot.send_message(entity=dst, message=message, parse_mode='html')
