"use client";

import React from "react";
import { Send, User, Bot, Loader2 } from "lucide-react";
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
    { role: "bot", content: "Hi! I'm InstantConnect Assistant. How can I help you book your session today?" }
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

      // Handle suggested actions from the AI
      if (result.suggestedAction) {
        if (result.suggestedAction.type === 'showPaymentLink') {
          onShowPayment?.();
        } else if (result.suggestedAction.type === 'showBookingSlots') {
          onShowSlots?.(result.suggestedAction.data as any[]);
        }
      }
    } catch (error) {
      console.error("Chat error:", error);
      setMessages(prev => [...prev, { role: "bot", content: "Sorry, I encountered an error. Please try again." }]);
    } finally {
      setLoading(false);
    }
  };

  const assistantAvatar = PlaceHolderImages.find(img => img.id === "assistant-avatar")?.imageUrl;

  return (
    <Card className="w-full flex flex-col h-[500px] border-border bg-card/30 backdrop-blur-md">
      <div className="p-4 border-b border-border flex items-center gap-3">
        <Avatar className="w-8 h-8 border border-primary/50">
          <AvatarImage src={assistantAvatar} alt="Bot" />
          <AvatarFallback><Bot className="w-4 h-4" /></AvatarFallback>
        </Avatar>
        <div>
          <h3 className="text-sm font-semibold font-headline">InstantConnect Assistant</h3>
          <p className="text-[10px] text-primary flex items-center gap-1">
            <span className="w-1.5 h-1.5 rounded-full bg-primary animate-pulse" />
            Active 24/7
          </p>
        </div>
      </div>

      <ScrollArea className="flex-1 p-4" ref={scrollRef}>
        <div className="space-y-4">
          {messages.map((msg, idx) => (
            <div
              key={idx}
              className={cn(
                "flex items-start gap-2 max-w-[85%]",
                msg.role === "user" ? "ml-auto flex-row-reverse" : ""
              )}
            >
              <Avatar className="w-6 h-6 mt-1 flex-shrink-0">
                <AvatarFallback className={cn("text-[10px]", msg.role === "user" ? "bg-muted" : "bg-primary")}>
                  {msg.role === "user" ? <User className="w-3 h-3" /> : <Bot className="w-3 h-3 text-primary-foreground" />}
                </AvatarFallback>
              </Avatar>
              <div
                className={cn(
                  "p-3 rounded-2xl text-sm",
                  msg.role === "user" 
                    ? "bg-primary text-primary-foreground rounded-tr-none" 
                    : "bg-muted text-foreground rounded-tl-none"
                )}
              >
                {msg.content}
              </div>
            </div>
          ))}
          {loading && (
            <div className="flex items-start gap-2">
              <Avatar className="w-6 h-6 mt-1">
                <AvatarFallback className="bg-primary"><Bot className="w-3 h-3 text-primary-foreground" /></AvatarFallback>
              </Avatar>
              <div className="bg-muted p-3 rounded-2xl rounded-tl-none">
                <Loader2 className="w-4 h-4 animate-spin text-primary" />
              </div>
            </div>
          )}
        </div>
      </ScrollArea>

      <div className="p-4 border-t border-border flex gap-2">
        <Input
          placeholder="Type your question..."
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && handleSend()}
          className="bg-muted/50 border-border focus-visible:ring-primary"
        />
        <Button size="icon" onClick={handleSend} disabled={loading} className="bg-primary hover:bg-primary/90 text-primary-foreground">
          <Send className="w-4 h-4" />
        </Button>
      </div>
    </Card>
  );
}