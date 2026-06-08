
"use client";

import React from "react";
import { ShieldCheck, Heart, Flower2, Star, Zap, UserCheck, MessageCircle, Crown, ShieldAlert, Sparkles, CheckCircle2 } from "lucide-react";
import { BookingActions } from "@/components/booking-actions";
import { BookingSlots } from "@/components/booking-slots";
import { Chatbot } from "@/components/chatbot";
import { Badge } from "@/components/ui/badge";
import { Card } from "@/components/ui/card";

export default function Home() {
  const [paymentConfirmed, setPaymentConfirmed] = React.useState(false);

  return (
    <main className="min-h-screen bg-white flex flex-col items-center pb-12 selection:bg-primary selection:text-white">
      {/* Premium Header Banner */}
      <div className="w-full bg-slate-900 text-white py-3 px-4 text-center overflow-hidden border-b-4 border-primary">
        <p className="text-[10px] font-black tracking-[0.4em] uppercase flex items-center justify-center gap-3">
          <Crown className="w-4 h-4 text-accent animate-pulse" /> 
          OFFICIAL REAL MEET VIP BOOKING PORTAL 
          <Crown className="w-4 h-4 text-accent animate-pulse" />
        </p>
      </div>
      
      <div className="w-full max-w-md px-5 flex flex-col space-y-8 mt-6">
        {/* Brand Identity */}
        <header className="flex flex-col items-center space-y-5">
          <div className="relative">
            <div className="w-24 h-24 bg-primary rounded-[2.5rem] flex items-center justify-center text-white shadow-[0_25px_50px_-12px_rgba(37,99,235,0.5)] transform -rotate-3 border-4 border-white">
              <Flower2 className="w-14 h-14" />
            </div>
            <div className="absolute -bottom-2 -right-2 bg-secondary text-white p-2.5 rounded-full shadow-2xl border-4 border-white">
              <ShieldCheck className="w-5 h-5" />
            </div>
          </div>
          
          <div className="text-center space-y-1">
            <h1 className="text-4xl font-black font-headline tracking-tighter text-slate-900 leading-none">
              REAL MEET <span className="text-primary italic font-serif">PORTAL</span>
            </h1>
            <p className="text-[11px] font-bold text-slate-400 tracking-[0.3em] uppercase">
              PREMIUM WELLNESS • CERTIFIED STAFF
            </p>
          </div>

          <div className="flex gap-2">
            <div className="flex items-center gap-1.5 bg-slate-100 px-4 py-2 rounded-2xl">
              <Star className="w-3 h-3 text-accent fill-accent" />
              <span className="text-[10px] font-black text-slate-600">4.9/5 RATING</span>
            </div>
            <div className="flex items-center gap-1.5 bg-secondary/10 px-4 py-2 rounded-2xl">
              <CheckCircle2 className="w-3 h-3 text-secondary" />
              <span className="text-[10px] font-black text-secondary uppercase">VERIFIED SERVICE</span>
            </div>
          </div>
        </header>

        {/* Main Service Card */}
        <Card className="border-none shadow-[0_50px_100px_-20px_rgba(0,0,0,0.15)] rounded-[3rem] bg-slate-50 overflow-hidden relative border-t-8 border-primary">
          <div className="p-8 space-y-8 relative z-10">
            <div className="space-y-4">
              <div className="flex items-center gap-2">
                <Sparkles className="w-6 h-6 text-accent" />
                <h2 className="text-2xl font-black font-headline leading-tight text-slate-900">
                  Elite Spa Experience
                </h2>
              </div>
              <p className="text-sm text-slate-500 leading-relaxed font-bold">
                Book premium real meet sessions with our highly trained <span className="text-primary">female staff</span> in your city.
              </p>
              
              {!paymentConfirmed && (
                <div className="bg-white border-2 border-slate-100 p-5 rounded-[2rem] space-y-3 shadow-sm">
                   <div className="flex items-start gap-3">
                      <ShieldAlert className="w-6 h-6 text-primary shrink-0 mt-0.5" />
                      <div className="space-y-1">
                        <p className="text-[12px] text-slate-900 font-black leading-tight uppercase">Security Verification Required</p>
                        <p className="text-[11px] text-slate-500 font-bold leading-tight">
                          Pay ₹49 one-time portal fee to unlock direct calls and WhatsApp verification.
                        </p>
                      </div>
                   </div>
                </div>
              )}
            </div>

            <BookingActions 
              paymentConfirmed={paymentConfirmed} 
              onPaymentSuccess={() => setPaymentConfirmed(true)} 
            />
          </div>
        </Card>

        {/* AI Portal Assistant - ENSURES USER GETS A REPLY */}
        <div className="space-y-4">
          <div className="flex items-center gap-3 px-2">
            <MessageCircle className="w-5 h-5 text-primary" />
            <h3 className="text-sm font-black text-slate-900 uppercase tracking-widest">Portal Assistant</h3>
            <Badge className="ml-auto bg-green-500 text-[9px] font-black">ONLINE</Badge>
          </div>
          <Chatbot paymentConfirmed={paymentConfirmed} />
        </div>

        {/* VIP Dashboard for Paid Users */}
        {paymentConfirmed && (
          <div className="space-y-8 animate-in fade-in slide-in-from-bottom-8 duration-700">
            <div className="flex items-center justify-between px-2">
              <h3 className="text-xl font-black font-headline text-slate-900">Elite Access Unlocked</h3>
              <Badge className="bg-secondary text-white border-none px-4 py-1.5 font-black text-[10px]">VERIFIED USER</Badge>
            </div>
            
            <BookingSlots 
              slots={[
                { id: '1', time: '17:00', display: 'Premium Wellness Session' },
                { id: '2', time: '19:30', display: 'Luxury Relaxation Therapy' },
                { id: '3', time: '21:00', display: 'VIP Real Meet Service' },
              ]} 
              onSelect={(slot) => console.log('Slot Reserved:', slot)}
            />
          </div>
        )}

        {/* Professional Trust Markers */}
        <div className="grid grid-cols-3 gap-4 pt-4">
           <div className="flex flex-col items-center gap-2 p-4 bg-slate-50 rounded-3xl">
             <ShieldCheck className="w-6 h-6 text-primary" />
             <span className="text-[9px] font-black text-slate-400 uppercase tracking-tighter">Secure</span>
           </div>
           <div className="flex flex-col items-center gap-2 p-4 bg-slate-50 rounded-3xl">
             <Heart className="w-6 h-6 text-primary" />
             <span className="text-[9px] font-black text-slate-400 uppercase tracking-tighter">Private</span>
           </div>
           <div className="flex flex-col items-center gap-2 p-4 bg-slate-50 rounded-3xl">
             <UserCheck className="w-6 h-6 text-primary" />
             <span className="text-[9px] font-black text-slate-400 uppercase tracking-tighter">Verified</span>
           </div>
        </div>

        <footer className="pt-8 pb-4 text-center">
          <p className="text-[10px] text-slate-300 font-black tracking-[0.5em] uppercase">
            © 2024 REAL MEET BOOKING SYSTEM
          </p>
        </footer>
      </div>
    </main>
  );
}
