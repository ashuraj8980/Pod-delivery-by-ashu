
"use client";

import React from "react";
import { Phone, MessageSquare, Send, Zap } from "lucide-react";
import { Button } from "@/components/ui/button";
import { PaymentDialog } from "@/components/payment-dialog";
import { notifyTelegram } from "@/app/actions/telegram";

interface BookingActionsProps {
  paymentConfirmed: boolean;
  onPaymentSuccess: () => void;
}

export function BookingActions({ paymentConfirmed, onPaymentSuccess }: BookingActionsProps) {
  const [showPayment, setShowPayment] = React.useState(false);
  const [pendingAction, setPendingAction] = React.useState<string | null>(null);

  const handleActionClick = async (action: string) => {
    if (action === "telegram") {
      // Telegram is FREE and DIRECT
      await notifyTelegram(`<b>🚀 FREE ENTRY:</b> User clicked "Book via Telegram". Moving to @Reallmeetbot.`);
      window.open("https://t.me/Reallmeetbot", "_blank");
      return;
    }

    // Call and WhatsApp require payment
    if (paymentConfirmed) {
      await notifyTelegram(`<b>✅ VERIFIED ACTION:</b> User initiated <b>${action.toUpperCase()}</b>.`);
      if (action === "whatsapp") {
        window.open("https://wa.me/+910000000000?text=I%20want%20to%20book%20a%20real%20meet%20session", "_blank");
      } else {
        alert(`Request Received! Our VIP manager will contact you for your session details.`);
      }
    } else {
      setPendingAction(action);
      setShowPayment(true);
    }
  };

  return (
    <div className="flex flex-col gap-4 w-full">
      {/* Primary Call Action - PAID */}
      <Button
        variant="default"
        onClick={() => handleActionClick("call")}
        className="h-20 text-xl font-black bg-primary hover:bg-primary/95 text-white shadow-2xl shadow-primary/30 rounded-[1.75rem] transition-all active:scale-95 group relative overflow-hidden border-b-4 border-primary-foreground/20"
      >
        <span className="flex items-center gap-3 relative z-10">
          <Phone className="w-6 h-6 fill-current" />
          BOOK VIA CALL
        </span>
        <div className="ml-auto bg-white/10 px-3 py-1.5 rounded-xl text-[10px] font-black uppercase tracking-widest border border-white/20 relative z-10">
          {paymentConfirmed ? "VERIFIED" : "₹49 FEE"}
        </div>
        <div className="absolute top-0 -left-[100%] w-full h-full bg-gradient-to-r from-transparent via-white/10 to-transparent group-hover:left-[100%] transition-all duration-1000" />
      </Button>

      <div className="grid grid-cols-1 gap-4">
        {/* Telegram VIP Action - FREE */}
        <Button
          variant="secondary"
          onClick={() => handleActionClick("telegram")}
          className="h-20 font-black bg-secondary hover:bg-secondary/95 text-white rounded-[1.75rem] flex items-center justify-between px-8 text-lg uppercase tracking-tight shadow-2xl shadow-secondary/30 transition-all border-b-4 border-black/10 animate-cta-pulse"
        >
          <div className="flex items-center gap-3">
            <Send className="w-6 h-6 fill-current" />
            <span>VIP TELEGRAM BOT</span>
          </div>
          <div className="bg-white/20 px-3 py-1 rounded-lg text-[10px] font-black">FREE ACCESS</div>
        </Button>

        {/* WhatsApp Action - PAID */}
        <Button
          variant="outline"
          onClick={() => handleActionClick("whatsapp")}
          className="h-16 font-black border-2 border-slate-200 hover:border-primary/20 hover:bg-primary/5 rounded-[1.5rem] flex items-center justify-center gap-2 text-xs uppercase tracking-widest transition-all"
        >
          <MessageSquare className="w-4 h-4 text-primary" />
          WHATSAPP (₹49)
        </Button>
      </div>

      <PaymentDialog 
        open={showPayment} 
        onOpenChange={setShowPayment} 
        onSuccess={() => {
          onPaymentSuccess();
          setShowPayment(false);
        }}
      />
    </div>
  );
}
