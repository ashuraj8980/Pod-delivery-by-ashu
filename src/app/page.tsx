"use client";

import React from "react";
import { Sparkles, ShieldCheck, Clock, Heart, Flower2, Star } from "lucide-react";
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
    <main className="min-h-screen bg-background relative overflow-x-hidden flex flex-col items-center">
      {/* Visual background accents */}
      <div className="fixed top-0 left-0 w-full h-full pointer-events-none overflow-hidden">
        <div className="absolute -top-20 -left-20 w-96 h-96 bg-primary/10 rounded-full blur-[100px]" />
        <div className="absolute top-1/2 -right-20 w-80 h-80 bg-secondary/10 rounded-full blur-[80px]" />
        <div className="absolute -bottom-20 left-1/4 w-72 h-72 bg-accent/10 rounded-full blur-[90px]" />
      </div>

      <div className="w-full max-w-lg z-10 px-4 flex flex-col space-y-8 py-10">
        {/* Navigation / Brand */}
        <header className="flex flex-col items-center space-y-4 animate-in fade-in slide-in-from-top-4 duration-700">
          <div className="flex items-center gap-2">
            <div className="w-10 h-10 bg-primary rounded-xl flex items-center justify-center text-primary-foreground shadow-lg">
              <Flower2 className="w-6 h-6" />
            </div>
            <span className="text-xl font-bold tracking-tight text-primary">REAL MEET</span>
          </div>
          <Badge variant="secondary" className="bg-secondary/10 text-secondary border-secondary/20 py-1 px-4 font-medium">
            <Star className="w-3 h-3 mr-2 fill-current" />
            Certified Female Staff Only
          </Badge>
          <div className="text-center space-y-2">
            <h1 className="text-4xl md:text-5xl font-bold font-headline tracking-tight leading-tight">
              Spa Booking <span className="text-primary italic">Portal</span>
            </h1>
            <p className="text-muted-foreground text-base max-w-[320px] mx-auto">
              Luxury real meet services for premium relaxation and professional wellness.
            </p>
          </div>
        </header>

        {/* Action Center */}
        <section className="bg-card glass-effect rounded-3xl p-6 shadow-2xl space-y-6">
          <div className="flex items-center justify-between px-2">
            <h2 className="text-sm font-semibold uppercase tracking-widest text-muted-foreground">Quick Booking</h2>
            <div className="flex gap-1">
              <div className="w-2 h-2 rounded-full bg-green-500 animate-pulse" />
              <span className="text-[10px] font-medium text-green-600">ONLINE</span>
            </div>
          </div>
          
          <BookingActions 
            paymentConfirmed={paymentConfirmed} 
            onPaymentSuccess={() => {
              setPaymentConfirmed(true);
              setShowSlots(true);
            }} 
          />
          
          <div className="grid grid-cols-3 gap-2 py-4">
            <div className="flex flex-col items-center space-y-1">
              <div className="p-2 bg-primary/5 rounded-full text-primary"><ShieldCheck className="w-5 h-5" /></div>
              <p className="text-[10px] font-semibold text-muted-foreground">VERIFIED</p>
            </div>
            <div className="flex flex-col items-center space-y-1">
              <div className="p-2 bg-secondary/5 rounded-full text-secondary"><Heart className="w-5 h-5" /></div>
              <p className="text-[10px] font-semibold text-muted-foreground">TRUSTED</p>
            </div>
            <div className="flex flex-col items-center space-y-1">
              <div className="p-2 bg-accent/5 rounded-full text-accent-foreground"><Clock className="w-5 h-5" /></div>
              <p className="text-[10px] font-semibold text-muted-foreground">24/7 LIVE</p>
            </div>
          </div>
        </section>

        <Separator className="opacity-50" />

        {/* Dynamic Booking/Chat Content */}
        <section className="space-y-6">
          {paymentConfirmed && showSlots && (
            <div className="animate-in fade-in slide-in-from-bottom-6 duration-500">
              <div className="flex items-center gap-2 mb-4">
                <div className="w-1 h-6 bg-secondary rounded-full" />
                <h2 className="text-xl font-headline font-bold">Select Your Slot</h2>
              </div>
              <BookingSlots 
                slots={availableSlots.length > 0 ? availableSlots : [
                  { id: '1', time: new Date().toISOString(), display: 'Evening, 5:00 PM - Spa Room 1' },
                  { id: '2', time: new Date().toISOString(), display: 'Evening, 7:30 PM - Spa Room 3' },
                  { id: '3', time: new Date().toISOString(), display: 'Tomorrow, 2:00 PM - VIP Suite' },
                ]} 
                onSelect={(slot) => console.log('Booking:', slot)}
              />
            </div>
          )}

          <div className="animate-in fade-in slide-in-from-bottom-8 duration-700">
            <div className="mb-4 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Sparkles className="w-5 h-5 text-accent-foreground" />
                <h2 className="text-xl font-headline font-bold">Live Support</h2>
              </div>
              {paymentConfirmed && (
                <Badge className="bg-secondary text-secondary-foreground font-bold">VIP UNLOCKED</Badge>
              )}
            </div>
            <Chatbot 
              paymentConfirmed={paymentConfirmed} 
              onShowPayment={() => {}}
              onShowSlots={(slots) => {
                setAvailableSlots(slots);
                setShowSlots(true);
              }}
            />
          </div>
        </section>

        <footer className="py-12 text-center">
          <p className="text-xs text-muted-foreground font-medium">© 2024 Real Meet Booking Portal - Professional Wellness</p>
          <div className="flex justify-center gap-6 mt-4 text-xs font-semibold text-primary/60">
            <a href="#" className="hover:text-primary underline decoration-primary/20">Privacy</a>
            <a href="#" className="hover:text-primary underline decoration-primary/20">Staff Verification</a>
            <a href="#" className="hover:text-primary underline decoration-primary/20">Safety Guide</a>
          </div>
        </footer>
      </div>
    </main>
  );
}