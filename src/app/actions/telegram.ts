
'use server';

/**
 * @fileOverview Server actions for Telegram Bot API communication.
 * Handles owner notifications, user replies, and webhook setup.
 */

const TELEGRAM_BOT_TOKEN = '8708245394:AAFtGFpXteDWcam_uNL-gV808tONgDDM8lc';
const OWNER_CHAT_ID = '8720928231';

/**
 * Notifies the owner (you) about portal activities.
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
    console.error('Notify Owner Error:', error);
    return { success: false };
  }
}

/**
 * Sends an AI response back to a specific user on Telegram.
 */
export async function sendMessageToUser(chatId: number | string, text: string) {
  try {
    const response = await fetch(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        chat_id: chatId,
        text: text,
        parse_mode: 'HTML',
      }),
    });
    const data = await response.json();
    return { success: data.ok };
  } catch (error) {
    console.error('Send Message Error:', error);
    return { success: false };
  }
}

/**
 * Action to set the webhook for the bot. 
 * This connects your Telegram Bot to your website's AI Brain.
 */
export async function setupBotWebhook(publicUrl: string) {
  // Remove any trailing slashes from the URL
  const cleanUrl = publicUrl.replace(/\/$/, "");
  const webhookUrl = `${cleanUrl}/api/telegram-webhook`;
  
  try {
    const response = await fetch(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/setWebhook?url=${webhookUrl}`);
    const data = await response.json();
    return { success: data.ok, description: data.description };
  } catch (error) {
    console.error('Webhook Setup Error:', error);
    return { success: false, error: 'Failed to connect bot' };
  }
}
