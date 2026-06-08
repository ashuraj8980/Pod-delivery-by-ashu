
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

    // Telegram sends the message inside body.message
    if (body.message && body.message.text) {
      const chatId = body.message.chat.id;
      const userText = body.message.text;
      const firstName = body.message.from.first_name || 'Sir';
      const username = body.message.from.username ? `@${body.message.from.username}` : firstName;

      // 1. Log the activity for the owner (you)
      // We don't await this to keep the response fast for the user
      notifyOwner(`<b>💬 Message from ${firstName} (${username})</b>\nText: ${userText}`);

      // 2. Get AI Response in Hinglish
      const aiResponse = await autonomousBookingAssistant({
        message: userText,
        conversationHistory: `User Name: ${firstName}`, // Basic context
        paymentConfirmed: false, // Webhook users are usually fresh inquiries
      });

      // 3. Send AI response back to the user on Telegram
      await sendMessageToUser(chatId, aiResponse.response);
    }

    // Always return 200 OK to Telegram so it doesn't keep retrying
    return NextResponse.json({ ok: true });
  } catch (error) {
    console.error('Webhook Error:', error);
    // Even on error, return 200 to Telegram unless it's a critical infrastructure failure
    return NextResponse.json({ ok: true });
  }
}
