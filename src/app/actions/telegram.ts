
'use server';

/**
 * @fileOverview Server actions for Telegram Bot API communication.
 * Handles both owner notifications and direct user replies.
 */

const TELEGRAM_BOT_TOKEN = '8708245394:AAFtGFpXteDWcam_uNL-gV808tONgDDM8lc';
const OWNER_CHAT_ID = '8720928231';

/**
 * Notifies the owner about portal activities.
 */
export async function notifyOwner(message: string) {
  try {
    await fetch(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        chat_id: OWNER_CHAT_ID,
        text: message,
        parse_mode: 'HTML',
      }),
    });
    return { success: true };
  } catch (error) {
    return { success: false };
  }
}

/**
 * Sends an AI response back to a specific user on Telegram.
 */
export async function sendMessageToUser(chatId: number | string, text: string) {
  try {
    await fetch(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        chat_id: chatId,
        text: text,
        parse_mode: 'HTML',
      }),
    });
    return { success: true };
  } catch (error) {
    return { success: false };
  }
}

/**
 * Action to set the webhook for the bot. 
 * Note: This needs a public URL to work (e.g. from App Hosting).
 */
export async function setupBotWebhook(publicUrl: string) {
  const webhookUrl = `${publicUrl}/api/telegram-webhook`;
  try {
    const response = await fetch(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/setWebhook?url=${webhookUrl}`);
    const data = await response.json();
    return { success: data.ok, description: data.description };
  } catch (error) {
    return { success: false, error: 'Failed to set webhook' };
  }
}
