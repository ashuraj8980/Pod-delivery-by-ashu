
'use server';
/**
 * @fileOverview A professional VIP Concierge for the Real Meet Booking Portal.
 * Handles customer inquiries immediately in Hinglish for a natural feel.
 */

import {ai} from '@/ai/genkit';
import {z} from 'genkit';

const AutonomousBookingAssistantInputSchema = z.object({
  message: z.string().describe('The user\'s current message to the bot.'),
  conversationHistory: z.string().describe('The previous turns of conversation.'),
  paymentConfirmed: z.boolean().describe('True if the user has paid the 49rs portal fee.'),
});
export type AutonomousBookingAssistantInput = z.infer<typeof AutonomousBookingAssistantInputSchema>;

const AutonomousBookingAssistantOutputSchema = z.object({
  response: z.string().describe('The AI\'s professional and helpful response in Hinglish.'),
  suggestedAction: z.discriminatedUnion('type', [
    z.object({
      type: z.literal('showPaymentLink'),
    }),
    z.object({
      type: z.literal('goToTelegram'),
    }),
    z.object({
      type: z.literal('none'),
    }),
  ]).optional(),
});
export type AutonomousBookingAssistantOutput = z.infer<typeof AutonomousBookingAssistantOutputSchema>;

const autonomousBookingAssistantPrompt = ai.definePrompt({
  name: 'autonomousBookingAssistantPrompt',
  input: { schema: AutonomousBookingAssistantInputSchema },
  output: { schema: AutonomousBookingAssistantOutputSchema },
  prompt: `You are the "VIP Concierge" for the Real Meet Booking Portal. 
Your goal is to provide IMMEDIATE, helpful, and professional responses in HINGLISH (Mix of Hindi and English) so local customers feel valued.

Context:
User message: {{{message}}}
Conversation History: {{{conversationHistory}}}
User has paid ₹49: {{{paymentConfirmed}}}

Guidelines:
1. Tone: Extremely professional yet friendly. Use Hinglish like "Ji sir, hamari female staff certified hain" or "Verification ke liye ₹49 mandatory hain".
2. Terms to use: "Wellness Session", "Certified Female Staff", "Elite Service", "Real Meet".
3. Handling No-Response Concerns: If user says Telegram pe reply nahi aa raha, say: "Don't worry sir, main yahan instant assistance ke liye hoon. Telegram bot coordination ke liye hai, par booking questions ke liye main aapki help karunga."
4. Fee Explanation: Explain ₹49 is a "Security & Verification Fee" for a safe environment.
5. Telegram Referral: Encourage them to use the "VIP Telegram Bot" for direct coordination as it is FREE.
`,
});

const autonomousBookingAssistantFlow = ai.defineFlow(
  {
    name: 'autonomousBookingAssistantFlow',
    inputSchema: AutonomousBookingAssistantInputSchema,
    outputSchema: AutonomousBookingAssistantOutputSchema,
  },
  async (input) => {
    const { output } = await autonomousBookingAssistantPrompt(input);
    return output!;
  }
);

export async function autonomousBookingAssistant(input: AutonomousBookingAssistantInput): Promise<AutonomousBookingAssistantOutput> {
  return autonomousBookingAssistantFlow(input);
}
