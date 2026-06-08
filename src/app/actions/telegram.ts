
'use server';

/**
 * @fileOverview Server action to notify the Telegram Bot when a user initiates a booking or payment.
 * Updated with user-provided credentials for @Reallmeetbot.
 */

const TELEGRAM_BOT_TOKEN = '8708245394:AAFtGFpXteDWcam_uNL-gV808tONgDDM8lc';
const TELEGRAM_CHAT_ID = '8720928231';

export async function notifyTelegram(message: string) {
  try {
    const response = await fetch(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        chat_id: TELEGRAM_CHAT_ID,
        text: message,
        parse_mode: 'HTML',
      }),
    });

    const data = await response.json();
    return { success: data.ok };
  } catch (error) {
    console.error('Telegram notification error:', error);
    return { success: false, error: 'Failed to send' };
  }
}
