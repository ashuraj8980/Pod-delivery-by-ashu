
"use client";

import React from "react";
import { ShieldCheck, Clock, Heart, Flower2, Star, Zap, UserCheck, MessageCircle } from "lucide-react";
import { BookingActions } from "@/components/booking-actions";
import { BookingSlots } from "@/components/booking-slots";
import { Separator } from "@/components/ui/separator";
import { Badge } from "@/components/ui/badge";
import { Card } from "@/components/ui/card";

export default function Home() {
  const [paymentConfirmed, setPaymentConfirmed] = React.useState(false);
  const [showSlots, setShowSlots] = React.useState(false);

  return (
    <main className="min-h-screen bg-[#F8FAFC] flex flex-col items-center pb-20">
      {/* Premium Header Decoration */}
      <div className="w-full h-2 bg-gradient-to-r from-primary via-secondary to-accent" />
      
      <div className="w-full max-w-md px-5 flex flex-col space-y-6 pt-8">
        {/* Brand Identity */}
        <header className="flex flex-col items-center space-y-4 text-center">
          <div className="relative">
            <div className="w-16 h-16 bg-primary rounded-2xl flex items-center justify-center text-white shadow-xl shadow-primary/20 rotate-3">
              <Flower2 className="w-10 h-10" />
            </div>
            <div className="absolute -top-2 -right-2 bg-accent text-accent-foreground p-1 rounded-full shadow-lg">
              <Star className="w-4 h-4 fill-current" />
            </div>
          </div>
          
          <div className="space-y-1">
            <h1 className="text-3xl font-black font-headline tracking-tight text-slate-900 uppercase">
              Real Meet <span className="text-primary italic font-light">Portal</span>
            </h1>
            <p className="text-xs font-bold text-muted-foreground tracking-[0.2em] uppercase">
              Verified Female Spa Staff • 24/7 Availability
            </p>
          </div>

          <div className="flex gap-2">
            <Badge variant="outline" className="border-secondary/30 bg-secondary/5 text-secondary font-bold px-3 py-1">
              <ShieldCheck className="w-3 h-3 mr-1" /> VERIFIED
            </Badge>
            <Badge variant="outline" className="border-primary/30 bg-primary/5 text-primary font-bold px-3 py-1">
              <UserCheck className="w-3 h-3 mr-1" /> PREMIUM
            </Badge>
          </div>
        </header>

        {/* Hero Card */}
        <Card className="border-none shadow-2xl rounded-[2.5rem] bg-white overflow-hidden">
          <div className="p-8 space-y-6">
            <div className="space-y-2">
              <h2 className="text-xl font-bold font-headline leading-tight">
                Unlock Professional Spa Sessions & Verified Profiles
              </h2>
              <p className="text-sm text-slate-500 leading-relaxed font-medium">
                To maintain privacy and security, a one-time portal fee of <span className="text-primary font-bold">₹49</span> is required to view our certified staff and book sessions.
              </p>
            </div>

            <BookingActions 
              paymentConfirmed={paymentConfirmed} 
              onPaymentSuccess={() => {
                setPaymentConfirmed(true);
                setShowSlots(true);
              }} 
            />

            <div className="grid grid-cols-3 gap-4 pt-4 border-t border-slate-100">
              <div className="text-center">
                <div className="text-lg font-bold text-primary italic">100%</div>
                <div className="text-[9px] font-bold text-slate-400 uppercase tracking-tighter">Verified Staff</div>
              </div>
              <div className="text-center border-x border-slate-100 px-2">
                <div className="text-lg font-bold text-secondary italic">VIP</div>
                <div className="text-[9px] font-bold text-slate-400 uppercase tracking-tighter">Private Slots</div>
              </div>
              <div className="text-center">
                <div className="text-lg font-bold text-accent-foreground italic">Fast</div>
                <div className="text-[9px] font-bold text-slate-400 uppercase tracking-tighter">Quick Booking</div>
              </div>
            </div>
          </div>
        </Card>

        {/* Dynamic Content */}
        {paymentConfirmed ? (
          <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
            <div className="flex items-center justify-between">
              <h3 className="text-lg font-bold font-headline">Available Today</h3>
              <Badge className="bg-green-100 text-green-700 hover:bg-green-100 border-none">ACTIVE</Badge>
            </div>
            
            <BookingSlots 
              slots={[
                { id: '1', time: '17:00', display: 'Premium Session - Deluxe Room (5:00 PM)' },
                { id: '2', time: '19:30', display: 'Elite Wellness - VIP Suite (7:30 PM)' },
                { id: '3', time: '21:00', display: 'Relaxation Special - Private Cabin (9:00 PM)' },
              ]} 
              onSelect={(slot) => console.log('Booking Slot:', slot)}
            />

            <Card className="bg-secondary text-secondary-foreground p-6 rounded-3xl border-none shadow-xl flex items-center gap-4">
              <div className="bg-white/20 p-3 rounded-2xl">
                <MessageCircle className="w-6 h-6" />
              </div>
              <div className="flex-1">
                <p className="text-sm font-bold leading-tight">Connect with VIP Manager</p>
                <p className="text-xs opacity-80">Finalize your booking details on Telegram</p>
              </div>
              <button 
                onClick={() => window.open('https://t.me/RealMeetPortalBot', '_blank')}
                className="bg-white text-secondary font-black px-4 py-2 rounded-xl text-xs hover:scale-105 transition-transform"
              >
                OPEN BOT
              </button>
            </Card>
          </div>
        ) : (
          <div className="space-y-4">
            <h3 className="text-center text-[10px] font-black text-slate-400 uppercase tracking-[0.3em]">Our Commitment</h3>
            <div className="grid grid-cols-1 gap-3">
              <div className="flex items-center gap-4 bg-white p-4 rounded-2xl shadow-sm border border-slate-100">
                <div className="w-10 h-10 bg-blue-50 text-primary rounded-xl flex items-center justify-center shrink-0">
                  <ShieldCheck className="w-5 h-5" />
                </div>
                <div>
                  <p className="text-xs font-bold text-slate-900">Privacy Guaranteed</p>
                  <p className="text-[10px] text-slate-500 font-medium">All sessions are handled with 100% discretion.</p>
                </div>
              </div>
              <div className="flex items-center gap-4 bg-white p-4 rounded-2xl shadow-sm border border-slate-100">
                <div className="w-10 h-10 bg-green-50 text-secondary rounded-xl flex items-center justify-center shrink-0">
                  <Heart className="w-5 h-5" />
                </div>
                <div>
                  <p className="text-xs font-bold text-slate-900">Certified Professionals</p>
                  <p className="text-[10px] text-slate-500 font-medium">Only highly trained female staff are onboarded.</p>
                </div>
              </div>
            </div>
          </div>
        )}

        <footer className="pt-12 pb-6 text-center space-y-4">
          <div className="flex justify-center items-center gap-2 grayscale opacity-50 text-[10px] font-black text-slate-400">
            <Zap className="w-3 h-3 fill-current" /> SECURE SSL ENCRYPTION
          </div>
          <p className="text-[10px] text-slate-400 font-bold tracking-widest uppercase">
            © 2024 Real Meet Booking Portal
          </p>
        </footer>
      </div>
    </main>
  );
}
