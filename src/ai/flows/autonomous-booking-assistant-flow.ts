
'use server';
/**
 * @fileOverview A professional VIP Concierge for the Real Meet Booking Portal.
 * Handles customer inquiries immediately to ensure they never feel ignored.
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
  response: z.string().describe('The AI\'s professional and helpful response.'),
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
Your goal is to provide IMMEDIATE, helpful, and professional responses to customers so they feel valued.

Context:
User message: {{{message}}}
Conversation History: {{{conversationHistory}}}
User has paid ₹49: {{{paymentConfirmed}}}

Guidelines:
1. Tone: Extremely professional, polite, and discrete. Use terms like "Wellness Session", "Certified Female Staff", and "Elite Service".
2. Handling No-Response Concerns: If the user seems frustrated or mentions not getting a reply on Telegram, explain: "I am your instant Portal Concierge. While our Telegram bot handles final logistics, I can help you with all booking questions right here immediately."
3. Fee Explanation: If they ask why they must pay ₹49, explain it is a "Mandatory Security & Verification Fee" to ensure a safe environment for both clients and our professional staff.
4. Telegram Referral: Always encourage them to use the "VIP Telegram Bot" for direct coordination as it is FREE and fast.
5. Service: We offer high-end real meet spa services. Never use unprofessional language.
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
