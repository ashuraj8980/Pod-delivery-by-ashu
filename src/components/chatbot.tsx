
"use client";

import React from "react";
import { Send, Loader2, Crown } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Card } from "@/components/ui/card";
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
    { role: "bot", content: "Hello sir! Main aapka VIP Concierge hoon. Real Meet booking ke liye main aapki kaise help kar sakta hoon?" }
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
      setMessages(prev => [...prev, { role: "bot", content: "Sorry sir, connection mein thoda issue hai. Please upar diye gaye buttons use karein." }]);
    } finally {
      setLoading(false);
    }
  };

  return (
    <Card className="w-full flex flex-col h-[450px] border-none bg-slate-50 shadow-2xl rounded-[2.5rem] overflow-hidden">
      <div className="p-5 border-b border-slate-200 bg-white flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 bg-primary rounded-full flex items-center justify-center text-white shadow-lg">
            <Crown className="w-5 h-5" />
          </div>
          <div className="flex flex-col">
            <span className="text-[11px] font-black text-slate-900 uppercase tracking-widest">VIP Concierge</span>
            <span className="text-[9px] font-bold text-green-500">ONLINE • INSTANT REPLY</span>
          </div>
        </div>
      </div>

      <ScrollArea className="flex-1 p-5 bg-[#F8FAFC]" ref={scrollRef}>
        <div className="space-y-4">
          {messages.map((msg, idx) => (
            <div
              key={idx}
              className={cn(
                "flex items-start gap-2 max-w-[88%]",
                msg.role === "user" ? "ml-auto flex-row-reverse" : ""
              )}
            >
              <div
                className={cn(
                  "p-4 px-5 rounded-[1.75rem] text-[13px] font-semibold leading-relaxed shadow-sm",
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
                <span className="text-[10px] font-black text-slate-400 uppercase">Bot is typing...</span>
              </div>
            </div>
          )}
        </div>
      </ScrollArea>

      <div className="p-4 bg-white border-t border-slate-100 flex gap-2">
        <Input
          placeholder="Apna sawaal likhein..."
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && handleSend()}
          className="rounded-full bg-slate-50 border-transparent focus-visible:ring-primary h-14 text-sm font-bold px-7"
        />
        <Button 
          size="icon" 
          onClick={handleSend} 
          disabled={loading} 
          className="rounded-full w-14 h-14 bg-primary hover:bg-primary/90 text-white shadow-xl active:scale-90 transition-transform"
        >
          <Send className="w-5 h-5" />
        </Button>
      </div>
    </Card>
  );
}
