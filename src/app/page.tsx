
"use client";

import React from "react";
import { ShieldCheck, Heart, Flower2, Star, Zap, UserCheck, MessageCircle, Crown, ShieldAlert, Sparkles } from "lucide-react";
import { BookingActions } from "@/components/booking-actions";
import { BookingSlots } from "@/components/booking-slots";
import { Badge } from "@/components/ui/badge";
import { Card } from "@/components/ui/card";

export default function Home() {
  const [paymentConfirmed, setPaymentConfirmed] = React.useState(false);

  return (
    <main className="min-h-screen bg-white flex flex-col items-center pb-12 selection:bg-primary selection:text-white">
      {/* Top Professional Banner */}
      <div className="w-full bg-primary text-white py-2.5 px-4 text-center overflow-hidden border-b-4 border-primary-foreground/10">
        <p className="text-[10px] font-black tracking-[0.3em] uppercase flex items-center justify-center gap-3">
          <Crown className="w-3.5 h-3.5 text-accent" /> 
          Verified Real Meet Spa Booking Portal 
          <Crown className="w-3.5 h-3.5 text-accent" />
        </p>
      </div>
      
      <div className="w-full max-w-md px-6 flex flex-col space-y-8 mt-8">
        {/* Brand Header */}
        <header className="flex flex-col items-center space-y-6">
          <div className="relative">
            <div className="w-24 h-24 bg-primary rounded-[2.5rem] flex items-center justify-center text-white shadow-[0_20px_50px_rgba(37,99,235,0.3)] transform -rotate-3 border-4 border-white">
              <Flower2 className="w-14 h-14" />
            </div>
            <div className="absolute -bottom-2 -right-2 bg-secondary text-white p-2 rounded-full shadow-xl border-4 border-white">
              <ShieldCheck className="w-5 h-5" />
            </div>
          </div>
          
          <div className="text-center space-y-2">
            <h1 className="text-4xl font-black font-headline tracking-tighter text-slate-900 leading-none">
              REAL MEET <span className="text-primary italic font-serif">PORTAL</span>
            </h1>
            <p className="text-[11px] font-bold text-slate-400 tracking-[0.3em] uppercase">
              PREMIUM WELLNESS • PROFESSIONAL STAFF
            </p>
          </div>

          <div className="flex gap-2">
            <Badge variant="outline" className="border-secondary/20 bg-secondary/10 text-secondary font-black px-5 py-2 rounded-xl text-[10px]">
              100% SECURE
            </Badge>
            <Badge variant="outline" className="border-accent/20 bg-accent/10 text-accent-foreground font-black px-5 py-2 rounded-xl text-[10px]">
              CERTIFIED STAFF
            </Badge>
          </div>
        </header>

        {/* Main Content Card */}
        <Card className="border-none shadow-[0_40px_80px_-15px_rgba(0,0,0,0.1)] rounded-[3rem] bg-slate-50 overflow-hidden relative border-t-8 border-primary/10">
          <div className="p-8 space-y-8 relative z-10">
            <div className="space-y-4">
              <div className="flex items-center gap-2">
                <Sparkles className="w-5 h-5 text-accent" />
                <h2 className="text-2xl font-black font-headline leading-tight text-slate-900">
                  Premium Spa Sessions
                </h2>
              </div>
              <p className="text-sm text-slate-500 leading-relaxed font-bold">
                Experience high-end relaxation with our highly trained <span className="text-primary">female therapists</span>.
              </p>
              
              {!paymentConfirmed && (
                <div className="bg-white border-2 border-slate-100 p-5 rounded-[2rem] space-y-3 shadow-sm">
                   <div className="flex items-start gap-3">
                      <ShieldAlert className="w-6 h-6 text-primary shrink-0 mt-0.5" />
                      <p className="text-[12px] text-slate-700 font-black leading-tight">
                        Direct call and WhatsApp verification requires a ₹49 portal fee.
                      </p>
                   </div>
                   <div className="pt-2 border-t border-slate-100 flex items-center justify-between">
                      <span className="text-[10px] font-black text-secondary uppercase tracking-widest">Free Option Available</span>
                      <Zap className="w-4 h-4 text-secondary fill-secondary" />
                   </div>
                </div>
              )}
            </div>

            <BookingActions 
              paymentConfirmed={paymentConfirmed} 
              onPaymentSuccess={() => setPaymentConfirmed(true)} 
            />

            <div className="grid grid-cols-3 gap-6 pt-8 border-t border-slate-200">
              <div className="text-center">
                <div className="text-xl font-black text-primary italic">24/7</div>
                <div className="text-[9px] font-black text-slate-400 uppercase tracking-widest">Booking</div>
              </div>
              <div className="text-center border-x border-slate-200 px-4">
                <div className="text-xl font-black text-secondary italic">VIP</div>
                <div className="text-[9px] font-black text-slate-400 uppercase tracking-widest">Bot</div>
              </div>
              <div className="text-center">
                <div className="text-xl font-black text-accent-foreground italic">SAFE</div>
                <div className="text-[9px] font-black text-slate-400 uppercase tracking-widest">Private</div>
              </div>
            </div>
          </div>
        </Card>

        {/* Dashboard for Paid Users */}
        {paymentConfirmed && (
          <div className="space-y-8 animate-in fade-in slide-in-from-bottom-8 duration-700">
            <div className="flex items-center justify-between px-2">
              <h3 className="text-xl font-black font-headline text-slate-900">Elite Access</h3>
              <Badge className="bg-green-500 text-white border-none px-4 py-1.5 font-black text-[10px]">VERIFIED</Badge>
            </div>
            
            <BookingSlots 
              slots={[
                { id: '1', time: '17:00', display: 'Elite Wellness Session' },
                { id: '2', time: '19:30', display: 'Premium Spa Therapy' },
                { id: '3', time: '21:00', display: 'Luxury Real Meet' },
              ]} 
              onSelect={(slot) => console.log('Slot Reserved:', slot)}
            />
          </div>
        )}

        {/* Benefits Section */}
        <div className="space-y-6 pt-4">
          <div className="flex items-center justify-center gap-8">
             <div className="flex flex-col items-center gap-2">
               <ShieldCheck className="w-7 h-7 text-primary opacity-40" />
               <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Secure</span>
             </div>
             <div className="flex flex-col items-center gap-2">
               <Heart className="w-7 h-7 text-primary opacity-40" />
               <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Privacy</span>
             </div>
             <div className="flex flex-col items-center gap-2">
               <UserCheck className="w-7 h-7 text-primary opacity-40" />
               <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Verified</span>
             </div>
          </div>
          <p className="text-center text-[10px] text-slate-400 font-bold max-w-[200px] mx-auto leading-relaxed">
            All therapists are certified professionals. We value your relaxation and privacy above all.
          </p>
        </div>

        <footer className="pt-12 pb-8 text-center">
          <p className="text-[10px] text-slate-300 font-black tracking-[0.5em] uppercase">
            © 2024 Real Meet Booking Portal
          </p>
        </footer>
      </div>
    </main>
  );
}
