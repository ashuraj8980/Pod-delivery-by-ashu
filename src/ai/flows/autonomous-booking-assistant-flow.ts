'use server';
/**
 * @fileOverview An autonomous booking assistant for the Real Meet Booking Portal.
 * Guides users through spa service selection, payment, and scheduling with female staff.
 */

import {ai} from '@/ai/genkit';
import {z} from 'genkit';

// Input Schema
const AutonomousBookingAssistantInputSchema = z.object({
  message: z.string().describe('The user\'s current message to the bot.'),
  conversationHistory: z.string().describe('The previous turns of conversation.'),
  paymentConfirmed: z.boolean().describe('True if the user has paid the 49rs portal fee.'),
});
export type AutonomousBookingAssistantInput = z.infer<typeof AutonomousBookingAssistantInputSchema>;

// Output Schema
const AutonomousBookingAssistantOutputSchema = z.object({
  response: z.string().describe('The AI\'s response focusing on spa services and professional female staff.'),
  suggestedAction: z.discriminatedUnion('type', [
    z.object({
      type: z.literal('showPaymentLink'),
      data: z.object({
        amount: z.number(),
        currency: z.string(),
        link: z.string(),
      }),
    }),
    z.object({
      type: z.literal('showBookingSlots'),
      data: z.array(z.object({
        id: z.string(),
        time: z.string(),
        display: z.string(),
      })),
    }),
    z.object({
      type: z.literal('connectToTelegramBot'),
      data: z.object({
        botLink: z.string(),
      }),
    }),
    z.object({
      type: z.literal('none'),
    }),
  ]).optional(),
});
export type AutonomousBookingAssistantOutput = z.infer<typeof AutonomousBookingAssistantOutputSchema>;

// Tools
const getPaymentLink = ai.defineTool(
  {
    name: 'getPaymentLink',
    description: 'Get the payment link for the 49rs portal registration fee.',
    inputSchema: z.object({}),
    outputSchema: z.object({ link: z.string(), amount: z.number(), currency: z.string() }),
  },
  async () => ({ link: 'https://pay.realmeetportal.com/49rs', amount: 49, currency: 'INR' })
);

const getAvailableBookingSlots = ai.defineTool(
  {
    name: 'getAvailableBookingSlots',
    description: 'Get available real meet spa session slots with female staff.',
    inputSchema: z.object({}),
    outputSchema: z.array(z.object({ id: z.string(), time: z.string(), display: z.string() })),
  },
  async () => {
    const now = new Date();
    return Array.from({ length: 3 }).map((_, i) => {
      const time = new Date(now.getTime() + (i + 1) * 3600000 * 24);
      return {
        id: `slot_${i}`,
        time: time.toISOString(),
        display: `${time.toLocaleDateString('en-IN', { weekday: 'short', day: 'numeric', month: 'short' })} at ${time.getHours()}:00 PM`,
      };
    });
  }
);

const getTelegramBotLink = ai.defineTool(
  {
    name: 'getTelegramBotLink',
    description: 'Provide the VIP Telegram bot link for direct coordination.',
    inputSchema: z.object({}),
    outputSchema: z.object({ botLink: z.string() }),
  },
  async () => ({ botLink: 'https://t.me/RealMeetPortalBot' })
);

const autonomousBookingAssistantPrompt = ai.definePrompt({
  name: 'autonomousBookingAssistantPrompt',
  input: { schema: AutonomousBookingAssistantInputSchema },
  output: { schema: AutonomousBookingAssistantOutputSchema },
  tools: [getPaymentLink, getAvailableBookingSlots, getTelegramBotLink],
  prompt: `You are the Lead Coordinator for "Real Meet Booking Portal". 
We provide premium, professional real meet spa services featuring highly trained female staff.

User context:
History: {{{conversationHistory}}}
Current: {{{message}}}
Paid: {{{paymentConfirmed}}}

Guidelines:
1. Professional Tone: Be polite, respectful, and helpful. Use words like "Wellness", "Relaxation", "Certified Female Staff", and "Professional Real Meet".
2. Fee Requirement: If not paid (paymentConfirmed=false), clearly explain that a ₹49 portal security fee is mandatory to view verified staff profiles and book slots. Use getPaymentLink.
3. Post-Payment: If paid, welcome them to the VIP portal. Offer to show available slots (getAvailableBookingSlots) or connect them to our direct Telegram bot (getTelegramBotLink) for personalized choices.
4. Privacy: Assure users that all bookings are private and secure.
5. Mobile View: Remind them that the portal is optimized for mobile booking.
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