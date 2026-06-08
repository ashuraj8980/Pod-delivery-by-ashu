
"use client";

import React from "react";
import { Phone, MessageSquare, Send } from "lucide-react";
import { Button } from "@/components/ui/button";
import { PaymentDialog } from "@/components/payment-dialog";
import { notifyOwner } from "@/app/actions/telegram";

interface BookingActionsProps {
  paymentConfirmed: boolean;
  onPaymentSuccess: () => void;
}

export function BookingActions({ paymentConfirmed, onPaymentSuccess }: BookingActionsProps) {
  const [showPayment, setShowPayment] = React.useState(false);
  const [pendingAction, setPendingAction] = React.useState<string | null>(null);

  const handleActionClick = async (action: string) => {
    if (action === "telegram") {
      await notifyOwner(`<b>🚀 FREE ENTRY:</b> User clicked "Book via Telegram". Moving to @Reallmeetbot.`);
      window.open("https://t.me/Reallmeetbot", "_blank");
      return;
    }

    if (paymentConfirmed) {
      await notifyOwner(`<b>✅ VERIFIED ACTION:</b> User initiated <b>${action.toUpperCase()}</b>.`);
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
      <Button
        variant="default"
        onClick={() => handleActionClick("call")}
        className="h-20 text-xl font-black bg-primary hover:bg-primary/95 text-white shadow-2xl shadow-primary/30 rounded-[1.75rem] transition-all active:scale-95 group relative overflow-hidden border-b-4 border-primary-foreground/20"
      >
        <span className="flex items-center gap-3 relative z-10">
          <Phone className="w-6 h-6 fill-current" />
          VERIFIED CALL
        </span>
        <div className="ml-auto bg-white/10 px-3 py-1.5 rounded-xl text-[10px] font-black uppercase tracking-widest border border-white/20 relative z-10">
          {paymentConfirmed ? "UNLOCKED" : "₹49"}
        </div>
      </Button>

      <div className="grid grid-cols-1 gap-4">
        <Button
          variant="outline"
          onClick={() => handleActionClick("whatsapp")}
          className="h-16 font-black border-2 border-slate-200 hover:border-primary/20 hover:bg-primary/5 rounded-[1.5rem] flex items-center justify-center gap-2 text-xs uppercase tracking-widest transition-all"
        >
          <MessageSquare className="w-4 h-4 text-primary" />
          WHATSAPP VERIFICATION
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
