"use client";

import React from "react";
import { Sparkles, ShieldCheck, Clock, Users } from "lucide-react";
import { BookingActions } from "@/components/booking-actions";
import { Chatbot } from "@/components/chatbot";
import { BookingSlots } from "@/components/booking-slots";
import { Separator } from "@/components/ui/separator";
import { Badge } from "@/components/ui/badge";

export default function Home() {
  const [paymentConfirmed, setPaymentConfirmed] = React.useState(false);
  const [availableSlots, setAvailableSlots] = React.useState<any[]>([]);
  const [showSlots, setShowSlots] = React.useState(false);

  return (
    <main className="min-h-screen bg-background relative overflow-hidden flex flex-col items-center justify-center p-4">
      {/* Decorative background blur */}
      <div className="absolute top-[-10%] left-[-10%] w-[40%] h-[40%] bg-primary/20 blur-[120px] rounded-full pointer-events-none" />
      <div className="absolute bottom-[-10%] right-[-10%] w-[40%] h-[40%] bg-secondary/20 blur-[120px] rounded-full pointer-events-none" />

      <div className="w-full max-w-lg z-10 flex flex-col space-y-12 py-12 md:py-24">
        {/* Header Section */}
        <section className="text-center space-y-4 animate-in fade-in slide-in-from-top-4 duration-700">
          <Badge variant="outline" className="border-primary/30 text-primary py-1 px-3 bg-primary/5">
            <Sparkles className="w-3 h-3 mr-2" />
            24/7 Professional Connect
          </Badge>
          <h1 className="text-5xl md:text-6xl font-bold font-headline leading-tight">
            InstantConnect <span className="text-primary italic">Bot</span>
          </h1>
          <p className="text-muted-foreground text-lg max-w-md mx-auto leading-relaxed">
            Your gateway to professional real meet services. Book your slot with confidence and ease.
          </p>
        </section>

        {/* Action Hub */}
        <section className="space-y-6">
          <BookingActions 
            paymentConfirmed={paymentConfirmed} 
            onPaymentSuccess={() => {
              setPaymentConfirmed(true);
              setShowSlots(true);
            }} 
          />
          
          <div className="grid grid-cols-3 gap-4 pt-4 px-4 text-center">
            <div className="space-y-1">
              <div className="flex justify-center text-primary/60"><ShieldCheck className="w-5 h-5" /></div>
              <p className="text-[10px] uppercase tracking-wider text-muted-foreground">Secure</p>
            </div>
            <div className="space-y-1">
              <div className="flex justify-center text-primary/60"><Clock className="w-5 h-5" /></div>
              <p className="text-[10px] uppercase tracking-wider text-muted-foreground">Instant</p>
            </div>
            <div className="space-y-1">
              <div className="flex justify-center text-primary/60"><Users className="w-5 h-5" /></div>
              <p className="text-[10px] uppercase tracking-wider text-muted-foreground">Verified</p>
            </div>
          </div>
        </section>

        <Separator className="bg-white/5" />

        {/* Dynamic Content: Slots or Chat */}
        <section className="space-y-8 px-2">
          {paymentConfirmed && showSlots && (
            <div className="animate-in fade-in slide-in-from-bottom-4 duration-500">
              <BookingSlots 
                slots={availableSlots.length > 0 ? availableSlots : [
                  { id: '1', time: new Date().toISOString(), display: 'Today, 2:30 PM' },
                  { id: '2', time: new Date().toISOString(), display: 'Today, 4:00 PM' },
                  { id: '3', time: new Date().toISOString(), display: 'Tomorrow, 10:00 AM' },
                ]} 
                onSelect={(slot) => console.log('Selected:', slot)}
              />
            </div>
          )}

          <div className="animate-in fade-in slide-in-from-bottom-8 duration-700 delay-300">
            <div className="mb-4 flex items-center justify-between">
              <h2 className="text-xl font-headline font-semibold">Instant Support</h2>
              {paymentConfirmed && (
                <Badge className="bg-primary/20 text-primary border-none">VIP Access Unlocked</Badge>
              )}
            </div>
            <Chatbot 
              paymentConfirmed={paymentConfirmed} 
              onShowPayment={() => {/* This is handled by booking actions but can be triggered from AI */}}
              onShowSlots={(slots) => {
                setAvailableSlots(slots);
                setShowSlots(true);
              }}
            />
          </div>
        </section>

        {/* Footer info */}
        <footer className="text-center text-xs text-muted-foreground/40 pt-12 pb-8">
          <p>© 2024 InstantConnect. All rights reserved.</p>
          <div className="flex items-center justify-center gap-4 mt-2">
            <a href="#" className="hover:text-primary transition-colors">Privacy</a>
            <a href="#" className="hover:text-primary transition-colors">Terms</a>
            <a href="#" className="hover:text-primary transition-colors">Safety</a>
          </div>
        </footer>
      </div>
    </main>
  );
}