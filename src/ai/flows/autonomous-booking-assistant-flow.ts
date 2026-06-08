'use server';
/**
 * @fileOverview An autonomous booking assistant for Telegram that answers user questions,
 * guides through payment, and helps with booking or connecting to the service.
 *
 * - autonomousBookingAssistant - The main function to interact with the bot.
 * - AutonomousBookingAssistantInput - The input type for the autonomousBookingAssistant function.
 * - AutonomousBookingAssistantOutput - The return type for the autonomousBookingAssistant function.
 */

import {ai} from '@/ai/genkit';
import {z} from 'genkit';

// Input Schema
const AutonomousBookingAssistantInputSchema = z.object({
  message: z.string().describe('The user\'s current message to the bot.'),
  conversationHistory: z.string().describe('The previous turns of conversation, formatted as "User: ...\\nBot: ...\\n".'),
  paymentConfirmed: z.boolean().describe('True if the user has already paid the 49rs fee, false otherwise.'),
});
export type AutonomousBookingAssistantInput = z.infer<typeof AutonomousBookingAssistantInputSchema>;

// Output Schema
const AutonomousBookingAssistantOutputSchema = z.object({
  response: z.string().describe('The AI\'s textual response to the user.'),
  suggestedAction: z.discriminatedUnion('type', [
    z.object({
      type: z.literal('showPaymentLink'),
      data: z.object({
        amount: z.number().describe('The payment amount.'),
        currency: z.string().describe('The currency.'),
        link: z.string().url().describe('The URL for the payment gateway.'),
      }),
    }).describe('Suggests displaying a payment link to the user.'),
    z.object({
      type: z.literal('showBookingSlots'),
      data: z.array(z.object({
        id: z.string().describe('Unique ID for the slot.'),
        time: z.string().datetime().describe('The available slot time in ISO format.'),
        display: z.string().describe('Human-readable display of the slot time.'),
      })).describe('Suggests displaying available booking slots to the user.'),
    }).describe('Suggests displaying available booking slots to the user.'),
    z.object({
      type: z.literal('connectToTelegramBot'),
      data: z.object({
        botLink: z.string().url().describe('The URL to redirect the user to the private Telegram bot.'),
      }),
    }).describe('Suggests redirecting the user to a private Telegram bot.'),
    z.object({
      type: z.literal('none'),
    }).describe('No specific action is suggested; continue conversation.'),
  ]).optional().describe('An optional action for the client application to take based on the bot\'s response.'),
});
export type AutonomousBookingAssistantOutput = z.infer<typeof AutonomousBookingAssistantOutputSchema>;

// --- Tools Definition ---

// Tool to get the payment link
const getPaymentLink = ai.defineTool(
  {
    name: 'getPaymentLink',
    description: 'Provides the official payment link for the 49rs fee to access the service. Call this when the user needs to pay or asks about payment.',
    inputSchema: z.object({}), // No specific input needed for a fixed payment link
    outputSchema: z.object({
      link: z.string().url().describe('The URL for the payment gateway.'),
      amount: z.number().describe('The payment amount.'),
      currency: z.string().describe('The currency.'),
    }),
  },
  async () => {
    // In a real application, this would interact with a payment service to generate a unique link.
    // For this example, return a dummy link.
    return {
      link: 'https://example.com/pay/instantconnect-49rs',
      amount: 49,
      currency: 'INR',
    };
  }
);

// Tool to get available booking slots
const getAvailableBookingSlots = ai.defineTool(
  {
    name: 'getAvailableBookingSlots',
    description: 'Retrieves a list of currently available meeting slots for the service. Call this when the user is ready to book a slot or asks about availability.',
    inputSchema: z.object({}), // No specific input for general availability
    outputSchema: z.array(z.object({
      id: z.string().describe('Unique ID for the slot.'),
      time: z.string().datetime().describe('The available slot time in ISO format.'),
      display: z.string().describe('Human-readable display of the slot time.'),
    })),
  },
  async () => {
    // In a real application, this would fetch from a database or booking service.
    // For this example, return dummy slots.
    const now = new Date();
    const slots = [];
    for (let i = 0; i < 3; i++) {
      const futureTime = new Date(now.getTime() + (i + 1) * 24 * 60 * 60 * 1000); // Next 3 days
      slots.push({
        id: `slot_${i + 1}`,
        time: futureTime.toISOString(),
        display: futureTime.toLocaleString('en-IN', {
          weekday: 'short',
          month: 'short',
          day: 'numeric',
          hour: 'numeric',
          minute: 'numeric',
        }),
      });
    }
    return slots;
  }
);

// Tool to get the Telegram bot link
const getTelegramBotLink = ai.defineTool(
  {
    name: 'getTelegramBotLink',
    description: 'Provides the direct link to the private Telegram bot for automated assistance. Call this when the user expresses a desire to connect via Telegram.',
    inputSchema: z.object({}),
    outputSchema: z.object({
      botLink: z.string().url().describe('The URL to the Telegram bot.'),
    }),
  },
  async () => {
    // In a real application, this might be dynamically generated or configured.
    return {
      botLink: 'https://t.me/InstantConnectBot', // Dummy Telegram bot link
    };
  }
);

// --- Prompt Definition ---

const autonomousBookingAssistantPrompt = ai.definePrompt({
  name: 'autonomousBookingAssistantPrompt',
  input: { schema: AutonomousBookingAssistantInputSchema },
  output: { schema: AutonomousBookingAssistantOutputSchema },
  tools: [getPaymentLink, getAvailableBookingSlots, getTelegramBotLink],
  prompt: `You are InstantConnect Bot, a friendly and helpful 24/7 AI chatbot on Telegram designed to assist users with booking real meet services.
Your primary goal is to guide potential customers through the process, answer their questions, handle payment inquiries, and facilitate booking.

Current Conversation History:
{{{conversationHistory}}}

User: {{{message}}}

Instructions:
1. Respond to the user's message in a conversational and helpful manner.
2. If the user has not yet paid (paymentConfirmed is false):
    - Clearly state that a 49rs fee is required to access booking or direct contact with a representative.
    - If the user asks about payment or how to proceed, use the 'getPaymentLink' tool to provide the payment details and link.
    - Do not offer booking slots or direct contact until payment is confirmed.
3. If the user has paid (paymentConfirmed is true):
    - Acknowledge their payment and express readiness to assist further.
    - If the user asks about booking a slot, use the 'getAvailableBookingSlots' tool to retrieve and present available times.
    - If the user explicitly asks to connect via Telegram, use the 'getTelegramBotLink' tool to provide the bot's direct link.
    - If the user asks about connecting via call or WhatsApp, explain that once a slot is booked, contact details will be shared.
4. Always provide a clear 'response' in your output.
5. Populate the 'suggestedAction' field ONLY if a specific client-side action is strongly implied by your response and the tools you used.
    - If you used 'getPaymentLink', set 'suggestedAction.type' to 'showPaymentLink' and 'suggestedAction.data' to the tool's output.
    - If you used 'getAvailableBookingSlots', set 'suggestedAction.type' to 'showBookingSlots' and 'suggestedAction.data' to the tool's output.
    - If you used 'getTelegramBotLink', set 'suggestedAction.type' to 'connectToTelegramBot' and 'suggestedAction.data' to the tool's output.
    - Otherwise, if no specific client action is needed, omit 'suggestedAction' or set it to type 'none'.
6. Keep responses concise and to the point, guiding the user to the next logical step.
`,
});

// --- Flow Definition ---

const autonomousBookingAssistantFlow = ai.defineFlow(
  {
    name: 'autonomousBookingAssistantFlow',
    inputSchema: AutonomousBookingAssistantInputSchema,
    outputSchema: AutonomousBookingAssistantOutputSchema,
  },
  async (input) => {
    const { output } = await autonomousBookingAssistantPrompt(input);
    if (!output) {
      throw new Error('No output received from the prompt.');
    }
    return output;
  }
);

// --- Wrapper Function ---

export async function autonomousBookingAssistant(input: AutonomousBookingAssistantInput): Promise<AutonomousBookingAssistantOutput> {
  return autonomousBookingAssistantFlow(input);
}
