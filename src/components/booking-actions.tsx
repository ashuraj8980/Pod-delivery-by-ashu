
"use client";

import React from "react";
import { Phone, MessageSquare, Send, ArrowRight, Star } from "lucide-react";
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
      // Logic for after payment
      await notifyTelegram(`User clicked <b>${action.toUpperCase()}</b> after payment confirmation.`);
      
      if (action === "telegram") {
        window.open("https://t.me/RealMeetPortalBot", "_blank");
      } else {
        alert(`Request Sent! A manager will contact you on your registered number via ${action.toUpperCase()}.`);
      }
    } else {
      setPendingAction(action);
      setShowPayment(true);
    }
  };

  return (
    <div className="flex flex-col gap-3 w-full">
      <Button
        variant="default"
        onClick={() => handleActionClick("call")}
        className="h-16 text-lg font-black bg-primary hover:bg-primary/90 text-white shadow-xl shadow-primary/20 rounded-[1.25rem] transition-all active:scale-95 group relative overflow-hidden"
      >
        <div className="absolute top-0 left-0 w-1 h-full bg-white/20" />
        <span className="flex items-center gap-3">
          <Phone className="w-5 h-5 fill-current" />
          BOOK VIA CALL
        </span>
        <div className="ml-auto bg-white/20 px-2 py-1 rounded-md text-[9px] font-black uppercase tracking-widest">
          {paymentConfirmed ? "UNLOCKED" : "₹49 FEE"}
        </div>
      </Button>

      <div className="grid grid-cols-2 gap-3">
        <Button
          variant="outline"
          onClick={() => handleActionClick("whatsapp")}
          className="h-14 font-black border-2 border-slate-100 hover:border-primary/20 hover:bg-primary/5 rounded-[1.25rem] flex items-center justify-center gap-2 text-xs"
        >
          <MessageSquare className="w-4 h-4 text-primary" />
          WHATSAPP
        </Button>
        <Button
          variant="secondary"
          onClick={() => handleActionClick("telegram")}
          className="h-14 font-black bg-secondary hover:bg-secondary/90 text-white rounded-[1.25rem] flex items-center justify-center gap-2 text-xs shadow-lg shadow-secondary/20"
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
          if (pendingAction) {
            handleActionClick(pendingAction);
          }
        }}
      />
    </div>
  );
}
