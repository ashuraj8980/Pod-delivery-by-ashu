"use client";

import React from "react";
import { Phone, MessageSquare, Send, ArrowRight } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
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
        window.open("https://t.me/InstantConnectBot", "_blank");
      } else {
        // Handle Call/WhatsApp logic here
        alert(`${action} service unlocked! Redirecting...`);
      }
    } else {
      setPendingAction(action);
      setShowPayment(true);
    }
  };

  return (
    <div className="grid grid-cols-1 gap-4 w-full max-w-sm mx-auto p-4">
      <Button
        variant="default"
        size="lg"
        onClick={() => handleActionClick("call")}
        className="h-16 text-lg font-semibold flex items-center justify-between px-6 bg-primary hover:bg-primary/90 text-primary-foreground group transition-all duration-300 animate-cta-pulse"
      >
        <div className="flex items-center gap-3">
          <Phone className="w-6 h-6" />
          <span className="font-headline italic">Book via Call</span>
        </div>
        {!paymentConfirmed ? (
          <Badge variant="secondary" className="bg-white/20 text-xs font-normal">₹49 Fee</Badge>
        ) : (
          <ArrowRight className="w-5 h-5 group-hover:translate-x-1 transition-transform" />
        )}
      </Button>

      <Button
        variant="outline"
        size="lg"
        onClick={() => handleActionClick("whatsapp")}
        className="h-16 text-lg font-semibold flex items-center justify-between px-6 border-primary/30 hover:bg-primary/10 text-primary group transition-all duration-300"
      >
        <div className="flex items-center gap-3">
          <MessageSquare className="w-6 h-6" />
          <span className="font-headline italic">Book via WhatsApp</span>
        </div>
        {!paymentConfirmed ? (
          <Badge variant="outline" className="border-primary/30 text-xs font-normal">₹49 Fee</Badge>
        ) : (
          <ArrowRight className="w-5 h-5 group-hover:translate-x-1 transition-transform" />
        )}
      </Button>

      <Button
        variant="ghost"
        size="lg"
        onClick={() => handleActionClick("telegram")}
        className="h-16 text-lg font-semibold flex items-center justify-between px-6 text-muted-foreground hover:text-primary hover:bg-primary/5 group transition-all duration-300"
      >
        <div className="flex items-center gap-3">
          <Send className="w-6 h-6" />
          <span className="font-headline italic">Connect Telegram</span>
        </div>
        <ArrowRight className="w-5 h-5 group-hover:translate-x-1 transition-transform" />
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