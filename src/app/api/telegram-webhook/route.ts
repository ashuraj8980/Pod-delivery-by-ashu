
import { NextRequest, NextResponse } from 'next/server';
import { sendMessageToUser, notifyOwner } from '@/app/actions/telegram';
import { autonomousBookingAssistant } from '@/ai/flows/autonomous-booking-assistant-flow';

/**
 * @fileOverview Next.js API Route Handler that acts as a Telegram Webhook.
 * It receives user messages from Telegram, gets an AI response in Hinglish,
 * and sends it back to the user.
 */

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();

    // Check if it's a standard text message
    if (body.message && body.message.text) {
      const chatId = body.message.chat.id;
      const userText = body.message.text;
      const username = body.message.from.username || body.message.from.first_name || 'Guest';

      // 1. Log the activity for the owner
      await notifyOwner(`<b>💬 New Bot Message</b>\nUser: @${username}\nMessage: ${userText}`);

      // 2. Get AI Response in Hinglish
      // We pass paymentConfirmed as false by default for unknown users
      const aiResponse = await autonomousBookingAssistant({
        message: userText,
        conversationHistory: '', // We could implement session tracking here if needed
        paymentConfirmed: false,
      });

      // 3. Send AI response back to the user on Telegram
      await sendMessageToUser(chatId, aiResponse.response);
    }

    return NextResponse.json({ ok: true });
  } catch (error) {
    console.error('Webhook Error:', error);
    return NextResponse.json({ ok: false }, { status: 500 });
  }
}
