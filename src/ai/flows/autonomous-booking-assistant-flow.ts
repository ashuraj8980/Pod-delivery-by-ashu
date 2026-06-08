
'use server';
/**
 * @fileOverview A professional VIP Concierge for the Real Meet Booking Portal.
 * Specialized for Telegram interactions in Hinglish.
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
      type: z.literal('goToWebsite'),
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
Your goal is to handle Telegram inquiries for professional Spa & Wellness sessions.

Context:
User message: {{{message}}}
Conversation History: {{{conversationHistory}}}
User has paid ₹49: {{{paymentConfirmed}}}

Guidelines:
1. Language: Always reply in HINGLISH (Mix of Hindi and English).
2. Tone: Extremely professional, respectful, and elite. Use terms like "Ji sir", "Zaroor", "Certified Female Staff", "Security Verification".
3. Booking Process:
   - If they ask for staff photos or numbers, explain: "Sir, security aur privacy ke liye manager verification mandatory hai. Aapko portal pe ₹49 security fee pay karni hogi, uske baad hum direct WhatsApp ya Call pe certified profiles share karenge."
4. Services: "Real Meet" wellness sessions by certified therapists.
5. Goal: Help the user understand the quality of service and guide them to the website for payment if they want direct contact/photos.
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
