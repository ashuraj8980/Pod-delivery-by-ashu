
"use client";

import React from "react";
import { Send, User, Bot, Loader2, Sparkles, MessageCircle, Crown } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Card } from "@/components/ui/card";
import { Avatar, AvatarFallback, AvatarImage } from "@/components/ui/avatar";
import { autonomousBookingAssistant } from "@/ai/flows/autonomous-booking-assistant-flow";
import { cn } from "@/lib/utils";

interface Message {
  role: "user" | "bot";
  content: string;
}

interface ChatbotProps {
  paymentConfirmed: boolean;
}

export function Chatbot({ paymentConfirmed }: ChatbotProps) {
  const [messages, setMessages] = React.useState<Message[]>([
    { role: "bot", content: "Hello! I am your VIP Concierge. How can I assist you with your professional real meet booking today?" }
  ]);
  const [input, setInput] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  const scrollRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    if (scrollRef.current) {
      const scrollContainer = scrollRef.current.querySelector('[data-radix-scroll-area-viewport]');
      if (scrollContainer) {
        scrollContainer.scrollTo({ top: scrollContainer.scrollHeight, behavior: 'smooth' });
      }
    }
  }, [messages, loading]);

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
    } catch (error) {
      console.error("Chat error:", error);
      setMessages(prev => [...prev, { role: "bot", content: "I'm sorry, I'm experiencing a brief connection issue. Please use our direct buttons above." }]);
    } finally {
      setLoading(false);
    }
  };

  return (
    <Card className="w-full flex flex-col h-[400px] border-none bg-slate-50 shadow-inner rounded-[2.5rem] overflow-hidden">
      <div className="p-4 border-b border-slate-200 bg-white/50 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 bg-primary rounded-full flex items-center justify-center text-white">
            <Crown className="w-4 h-4" />
          </div>
          <span className="text-[10px] font-black text-slate-900 uppercase tracking-widest">Instant Concierge</span>
        </div>
        <div className="flex items-center gap-1.5">
          <div className="w-1.5 h-1.5 rounded-full bg-green-500 animate-pulse" />
          <span className="text-[9px] font-black text-green-600">LIVE</span>
        </div>
      </div>

      <ScrollArea className="flex-1 p-5" ref={scrollRef}>
        <div className="space-y-4">
          {messages.map((msg, idx) => (
            <div
              key={idx}
              className={cn(
                "flex items-start gap-2 max-w-[85%]",
                msg.role === "user" ? "ml-auto flex-row-reverse" : ""
              )}
            >
              <div
                className={cn(
                  "p-3.5 px-5 rounded-[1.5rem] text-[12px] font-bold leading-relaxed shadow-sm",
                  msg.role === "user" 
                    ? "bg-primary text-white rounded-tr-none" 
                    : "bg-white text-slate-700 border border-slate-100 rounded-tl-none"
                )}
              >
                {msg.content}
              </div>
            </div>
          ))}
          {loading && (
            <div className="flex items-start gap-2">
              <div className="bg-white border border-slate-100 p-3 px-5 rounded-[1.5rem] rounded-tl-none shadow-sm flex items-center gap-2">
                <Loader2 className="w-3 h-3 animate-spin text-primary" />
                <span className="text-[10px] font-black text-slate-400">TYPING...</span>
              </div>
            </div>
          )}
        </div>
      </ScrollArea>

      <div className="p-4 bg-white border-t border-slate-100 flex gap-2">
        <Input
          placeholder="Type your question..."
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && handleSend()}
          className="rounded-full bg-slate-50 border-transparent focus-visible:ring-primary h-12 text-sm font-bold px-6"
        />
        <Button size="icon" onClick={handleSend} disabled={loading} className="rounded-full w-12 h-12 bg-primary hover:bg-primary/90 text-white shadow-xl">
          <Send className="w-4 h-4" />
        </Button>
      </div>
    </Card>
  );
}
