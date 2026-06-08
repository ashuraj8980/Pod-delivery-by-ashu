"use client";

import React from "react";
import { Send, User, Bot, Loader2, Sparkles } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Card } from "@/components/ui/card";
import { Avatar, AvatarFallback, AvatarImage } from "@/components/ui/avatar";
import { autonomousBookingAssistant } from "@/ai/flows/autonomous-booking-assistant-flow";
import { PlaceHolderImages } from "@/lib/placeholder-images";
import { cn } from "@/lib/utils";

interface Message {
  role: "user" | "bot";
  content: string;
}

interface ChatbotProps {
  paymentConfirmed: boolean;
  onShowPayment?: () => void;
  onShowSlots?: (slots: any[]) => void;
}

export function Chatbot({ paymentConfirmed, onShowPayment, onShowSlots }: ChatbotProps) {
  const [messages, setMessages] = React.useState<Message[]>([
    { role: "bot", content: "Welcome to Real Meet Booking Portal! I'm your coordinator. How can I help you book a luxury spa session with our professional staff today?" }
  ]);
  const [input, setInput] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  const scrollRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    if (scrollRef.current) {
      scrollRef.current.scrollTo(0, scrollRef.current.scrollHeight);
    }
  }, [messages]);

  const handleSend = async () => {
    if (!input.trim() || loading) return;

    const userMsg = input.trim();
    setInput("");
    setMessages(prev => [...prev, { role: "user", content: userMsg }]);
    setLoading(true);

    try {
      const history = messages
        .map(m => `${m.role === "user" ? "User" : "Bot"}: ${m.content}`)
        .join("\n");

      const result = await autonomousBookingAssistant({
        message: userMsg,
        conversationHistory: history,
        paymentConfirmed
      });

      setMessages(prev => [...prev, { role: "bot", content: result.response }]);

      if (result.suggestedAction) {
        if (result.suggestedAction.type === 'showPaymentLink') {
          onShowPayment?.();
        } else if (result.suggestedAction.type === 'showBookingSlots') {
          onShowSlots?.(result.suggestedAction.data as any[]);
        }
      }
    } catch (error) {
      console.error("Chat error:", error);
      setMessages(prev => [...prev, { role: "bot", content: "Apologies, our coordination system is busy. Please try again or use direct booking buttons above." }]);
    } finally {
      setLoading(false);
    }
  };

  const assistantAvatar = PlaceHolderImages.find(img => img.id === "spa-therapist")?.imageUrl;

  return (
    <Card className="w-full flex flex-col h-[500px] border-border bg-card shadow-lg rounded-3xl overflow-hidden">
      <div className="p-4 border-b border-border bg-primary/5 flex items-center gap-3">
        <Avatar className="w-10 h-10 border-2 border-primary/20">
          <AvatarImage src={assistantAvatar} alt="Spa Coordinator" />
          <AvatarFallback className="bg-primary text-primary-foreground"><Bot className="w-5 h-5" /></AvatarFallback>
        </Avatar>
        <div>
          <h3 className="text-sm font-bold font-headline">Portal Coordinator</h3>
          <p className="text-[10px] text-green-600 font-bold flex items-center gap-1">
            <span className="w-1.5 h-1.5 rounded-full bg-green-500 animate-pulse" />
            READY TO ASSIST
          </p>
        </div>
      </div>

      <ScrollArea className="flex-1 p-4 bg-muted/20" ref={scrollRef}>
        <div className="space-y-4">
          {messages.map((msg, idx) => (
            <div
              key={idx}
              className={cn(
                "flex items-start gap-2 max-w-[90%]",
                msg.role === "user" ? "ml-auto flex-row-reverse" : ""
              )}
            >
              <div
                className={cn(
                  "p-3 px-4 rounded-2xl text-sm shadow-sm",
                  msg.role === "user" 
                    ? "bg-primary text-primary-foreground rounded-tr-none" 
                    : "bg-white text-foreground border border-border rounded-tl-none"
                )}
              >
                {msg.content}
              </div>
            </div>
          ))}
          {loading && (
            <div className="flex items-start gap-2">
              <div className="bg-white border border-border p-3 rounded-2xl rounded-tl-none shadow-sm">
                <Loader2 className="w-4 h-4 animate-spin text-primary" />
              </div>
            </div>
          )}
        </div>
      </ScrollArea>

      <div className="p-4 bg-white border-t border-border flex gap-2">
        <Input
          placeholder="Ask about spa services or staff..."
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && handleSend()}
          className="rounded-full bg-muted/50 border-transparent focus-visible:ring-primary h-11"
        />
        <Button size="icon" onClick={handleSend} disabled={loading} className="rounded-full w-11 h-11 bg-primary hover:bg-primary/90 text-primary-foreground shadow-lg">
          <Send className="w-4 h-4" />
        </Button>
      </div>
    </Card>
  );
}