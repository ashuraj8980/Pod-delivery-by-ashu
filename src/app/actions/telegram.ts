
'use server';

/**
 * @fileOverview Server action to notify the Telegram Bot when a user initiates a booking or payment.
 */

const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN || 'YOUR_BOT_TOKEN';
const TELEGRAM_CHAT_ID = process.env.TELEGRAM_CHAT_ID || 'YOUR_CHAT_ID';

export async function notifyTelegram(message: string) {
  if (!TELEGRAM_BOT_TOKEN || TELEGRAM_BOT_TOKEN === 'YOUR_BOT_TOKEN') {
    console.warn('Telegram Bot Token not configured');
    return { success: false, error: 'Not configured' };
  }

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
