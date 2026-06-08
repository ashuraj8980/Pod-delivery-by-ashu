"use client";

import React from "react";
import { Phone, MessageSquare, Send, ArrowRight, UserCheck } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { PaymentDialog } from "@/components/payment-dialog";

interface BookingActionsProps {
  paymentConfirmed: boolean;
  onPaymentSuccess: () => void;
}

export function BookingActions({ paymentConfirmed, onPaymentSuccess }: BookingActionsProps) {
  const [showPayment, setShowPayment] = React.useState(false);
  const [pendingAction, setPendingAction] = React.useState<string | null>(null);

  const handleActionClick = (action: string) => {
    if (paymentConfirmed) {
      if (action === "telegram") {
        window.open("https://t.me/RealMeetPortalBot", "_blank");
      } else {
        alert(`${action} coordinates unlocked! Our manager will call you shortly to finalize details.`);
      }
    } else {
      setPendingAction(action);
      setShowPayment(true);
    }
  };

  return (
    <div className="grid grid-cols-1 gap-3 w-full">
      <Button
        variant="default"
        size="lg"
        onClick={() => handleActionClick("call")}
        className="h-14 text-base font-bold flex items-center justify-between px-6 bg-primary hover:bg-primary/90 text-primary-foreground group rounded-2xl animate-cta-pulse"
      >
        <div className="flex items-center gap-3">
          <Phone className="w-5 h-5" />
          <span className="font-headline italic">Book Session via Call</span>
        </div>
        {!paymentConfirmed ? (
          <Badge variant="secondary" className="bg-white/20 text-[10px]">₹49 FEE</Badge>
        ) : (
          <ArrowRight className="w-4 h-4 group-hover:translate-x-1 transition-transform" />
        )}
      </Button>

      <Button
        variant="outline"
        size="lg"
        onClick={() => handleActionClick("whatsapp")}
        className="h-14 text-base font-bold flex items-center justify-between px-6 border-primary/20 hover:bg-primary/5 text-primary group rounded-2xl"
      >
        <div className="flex items-center gap-3">
          <MessageSquare className="w-5 h-5" />
          <span className="font-headline italic">WhatsApp Coordination</span>
        </div>
        {!paymentConfirmed ? (
          <Badge variant="outline" className="border-primary/20 text-[10px]">₹49 FEE</Badge>
        ) : (
          <ArrowRight className="w-4 h-4 group-hover:translate-x-1 transition-transform" />
        )}
      </Button>

      <Button
        variant="secondary"
        size="lg"
        onClick={() => handleActionClick("telegram")}
        className="h-14 text-base font-bold flex items-center justify-between px-6 bg-secondary hover:bg-secondary/90 text-secondary-foreground group rounded-2xl"
      >
        <div className="flex items-center gap-3">
          <Send className="w-5 h-5" />
          <span className="font-headline italic">VIP Telegram Bot</span>
        </div>
        <UserCheck className="w-4 h-4 group-hover:scale-110 transition-transform" />
      </Button>

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