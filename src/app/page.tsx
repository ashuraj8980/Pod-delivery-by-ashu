
"use client";

import React from "react";
import { ShieldCheck, Heart, Flower2, Star, Zap, UserCheck, MessageCircle, Crown, ShieldAlert } from "lucide-react";
import { BookingActions } from "@/components/booking-actions";
import { BookingSlots } from "@/components/booking-slots";
import { Badge } from "@/components/ui/badge";
import { Card } from "@/components/ui/card";

export default function Home() {
  const [paymentConfirmed, setPaymentConfirmed] = React.useState(false);

  return (
    <main className="min-h-screen bg-slate-50 flex flex-col items-center pb-12 selection:bg-primary selection:text-white">
      {/* Top Banner */}
      <div className="w-full bg-slate-900 text-white py-2 px-4 text-center overflow-hidden">
        <p className="text-[10px] font-black tracking-[0.3em] uppercase animate-pulse flex items-center justify-center gap-2">
          <Crown className="w-3 h-3 text-accent" /> 
          Verified Real Meet Spa Booking Portal 
          <Crown className="w-3 h-3 text-accent" />
        </p>
      </div>
      
      <div className="w-full max-w-md px-6 flex flex-col space-y-8 mt-6">
        {/* Brand Header */}
        <header className="flex flex-col items-center space-y-5">
          <div className="relative">
            <div className="w-20 h-20 bg-primary rounded-[2rem] flex items-center justify-center text-white shadow-2xl shadow-primary/30 transform -rotate-3 border-4 border-white">
              <Flower2 className="w-12 h-12" />
            </div>
            <div className="absolute -bottom-2 -right-2 bg-secondary text-white p-1.5 rounded-full shadow-lg border-2 border-white">
              <ShieldCheck className="w-4 h-4" />
            </div>
          </div>
          
          <div className="text-center space-y-1">
            <h1 className="text-4xl font-black font-headline tracking-tighter text-slate-900 leading-none">
              REAL MEET <span className="text-primary italic font-serif">PORTAL</span>
            </h1>
            <p className="text-[11px] font-bold text-slate-400 tracking-[0.25em] uppercase">
              LUXURY WELLNESS • CERTIFIED STAFF
            </p>
          </div>

          <div className="flex gap-2">
            <Badge variant="outline" className="border-secondary/20 bg-secondary/10 text-secondary font-black px-4 py-1.5 rounded-xl">
              100% SECURE
            </Badge>
            <Badge variant="outline" className="border-accent/20 bg-accent/10 text-accent-foreground font-black px-4 py-1.5 rounded-xl">
              VIP ONLY
            </Badge>
          </div>
        </header>

        {/* Main Content Card */}
        <Card className="border-none shadow-[0_32px_64px_-12px_rgba(0,0,0,0.1)] rounded-[3rem] bg-white overflow-hidden relative">
          {/* Decorative element */}
          <div className="absolute top-0 right-0 w-32 h-32 bg-primary/5 rounded-full -mr-16 -mt-16 blur-3xl" />
          
          <div className="p-8 space-y-8 relative z-10">
            <div className="space-y-3">
              <h2 className="text-2xl font-black font-headline leading-tight text-slate-900">
                Premium Real Meet Services with Female Staff
              </h2>
              <p className="text-sm text-slate-500 leading-relaxed font-medium">
                Our portal features only <span className="text-secondary font-bold">verified profiles</span> of professional female spa therapists.
              </p>
              
              {!paymentConfirmed && (
                <div className="bg-slate-50 border border-slate-100 p-4 rounded-[1.5rem] flex items-start gap-3">
                  <ShieldAlert className="w-5 h-5 text-primary shrink-0 mt-0.5" />
                  <p className="text-[11px] text-slate-600 font-bold leading-relaxed">
                    A security verification fee of <span className="text-primary">₹49</span> is mandatory to filter spam and protect staff privacy.
                  </p>
                </div>
              )}
            </div>

            <BookingActions 
              paymentConfirmed={paymentConfirmed} 
              onPaymentSuccess={() => setPaymentConfirmed(true)} 
            />

            <div className="grid grid-cols-3 gap-6 pt-6 border-t border-slate-50">
              <div className="text-center">
                <div className="text-xl font-black text-primary italic">24/7</div>
                <div className="text-[9px] font-black text-slate-400 uppercase tracking-widest">Support</div>
              </div>
              <div className="text-center border-x border-slate-100 px-4">
                <div className="text-xl font-black text-secondary italic">GOLD</div>
                <div className="text-[9px] font-black text-slate-400 uppercase tracking-widest">Profiles</div>
              </div>
              <div className="text-center">
                <div className="text-xl font-black text-accent-foreground italic">SAFE</div>
                <div className="text-[9px] font-black text-slate-400 uppercase tracking-widest">Private</div>
              </div>
            </div>
          </div>
        </Card>

        {/* Post-Payment Layout */}
        {paymentConfirmed && (
          <div className="space-y-8 animate-in fade-in slide-in-from-bottom-8 duration-700">
            <div className="flex items-center justify-between px-2">
              <h3 className="text-xl font-black font-headline text-slate-900">VIP Dashboard</h3>
              <Badge className="bg-green-500 text-white border-none px-3 py-1 font-black animate-pulse">LIVE ACCESS</Badge>
            </div>
            
            <BookingSlots 
              slots={[
                { id: '1', time: '17:00', display: 'Elite Wellness Session (5:00 PM)' },
                { id: '2', time: '19:30', display: 'Premium Spa Therapy (7:30 PM)' },
                { id: '3', time: '21:00', display: 'Luxury Real Meet (9:00 PM)' },
              ]} 
              onSelect={(slot) => console.log('Booking Slot Selected:', slot)}
            />

            <Card className="bg-slate-900 text-white p-8 rounded-[2.5rem] border-none shadow-2xl relative overflow-hidden group">
              <div className="absolute top-0 right-0 w-32 h-32 bg-primary/20 rounded-full -mr-16 -mt-16 blur-2xl group-hover:bg-primary/30 transition-colors" />
              <div className="space-y-6 relative z-10">
                <div className="flex items-center gap-4">
                  <div className="bg-primary p-4 rounded-[1.25rem] shadow-lg shadow-primary/20">
                    <MessageCircle className="w-8 h-8" />
                  </div>
                  <div>
                    <h4 className="text-lg font-black tracking-tight">VIP COORDINATOR</h4>
                    <p className="text-xs text-slate-400 font-bold">Connected to @Reallmeetbot</p>
                  </div>
                </div>
                <p className="text-sm text-slate-300 font-medium leading-relaxed">
                  Our manager is ready to handle your staff preferences and location details directly on Telegram.
                </p>
                <button 
                  onClick={() => window.open('https://t.me/Reallmeetbot', '_blank')}
                  className="w-full bg-white text-slate-900 font-black h-14 rounded-2xl text-sm flex items-center justify-center gap-2 hover:bg-slate-100 transition-all shadow-xl active:scale-95"
                >
                  CHAT ON TELEGRAM <Zap className="w-4 h-4 fill-primary text-primary" />
                </button>
              </div>
            </Card>
          </div>
        )}

        {/* Privacy Section */}
        {!paymentConfirmed && (
          <div className="space-y-6 pt-4">
            <div className="flex items-center justify-center gap-6">
               <div className="flex flex-col items-center gap-2 opacity-60">
                 <ShieldCheck className="w-6 h-6 text-slate-400" />
                 <span className="text-[9px] font-black text-slate-400 uppercase">Secure</span>
               </div>
               <div className="flex flex-col items-center gap-2 opacity-60">
                 <Heart className="w-6 h-6 text-slate-400" />
                 <span className="text-[9px] font-black text-slate-400 uppercase">Professional</span>
               </div>
               <div className="flex flex-col items-center gap-2 opacity-60">
                 <UserCheck className="w-6 h-6 text-slate-400" />
                 <span className="text-[9px] font-black text-slate-400 uppercase">Verified</span>
               </div>
            </div>
          </div>
        )}

        <footer className="pt-12 pb-8 text-center space-y-4">
          <p className="text-[10px] text-slate-400 font-black tracking-[0.4em] uppercase">
            © 2024 Real Meet Booking Portal
          </p>
        </footer>
      </div>
    </main>
  );
}
