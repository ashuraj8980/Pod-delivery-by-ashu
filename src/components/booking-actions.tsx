
"use client";

import React from "react";
import { Phone, MessageSquare, Send } from "lucide-react";
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
    if (paymentConfirmed) {
      // User is verified, notify telegram
      await notifyTelegram(`User initiated <b>${action.toUpperCase()}</b>. Session unlocking...`);
      
      if (action === "telegram") {
        window.open("https://t.me/Reallmeetbot", "_blank");
      } else {
        alert(`Request Received! Our VIP manager will contact you on WhatsApp/Call shortly.`);
      }
    } else {
      setPendingAction(action);
      setShowPayment(true);
    }
  };

  return (
    <div className="flex flex-col gap-4 w-full">
      {/* Primary Call Action */}
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
          {paymentConfirmed ? "VERIFIED" : "₹49"}
        </div>
        <div className="absolute top-0 -left-[100%] w-full h-full bg-gradient-to-r from-transparent via-white/10 to-transparent group-hover:left-[100%] transition-all duration-1000" />
      </Button>

      <div className="grid grid-cols-2 gap-4">
        <Button
          variant="outline"
          onClick={() => handleActionClick("whatsapp")}
          className="h-16 font-black border-2 border-slate-100 hover:border-primary/20 hover:bg-primary/5 rounded-[1.5rem] flex items-center justify-center gap-2 text-xs uppercase tracking-widest transition-all"
        >
          <MessageSquare className="w-4 h-4 text-primary" />
          WHATSAPP
        </Button>
        <Button
          variant="secondary"
          onClick={() => handleActionClick("telegram")}
          className="h-16 font-black bg-secondary hover:bg-secondary/95 text-white rounded-[1.5rem] flex items-center justify-center gap-2 text-xs uppercase tracking-widest shadow-xl shadow-secondary/20 transition-all border-b-4 border-black/10"
        >
          <Send className="w-4 h-4 fill-current" />
          VIP BOT
        </Button>
      </div>

      <PaymentDialog 
        open={showPayment} 
        onOpenChange={setShowPayment} 
        onSuccess={() => {
          onPaymentSuccess();
          setShowPayment(false);
          // If there was an action pending, we don't trigger it immediately to avoid popups during render.
          // The user can now click the unlocked buttons.
        }}
      />
    </div>
  );
}
