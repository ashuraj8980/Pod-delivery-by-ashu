
"use client";

import React from "react";
import { ShieldCheck, Heart, Flower2, Star, Zap, UserCheck, MessageCircle, Crown, ShieldAlert, Sparkles, CheckCircle2, ArrowRight, Settings } from "lucide-react";
import { BookingActions } from "@/components/booking-actions";
import { Badge } from "@/components/ui/badge";
import { Card } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { setupBotWebhook } from "@/app/actions/telegram";
import { toast } from "@/hooks/use-toast";

export default function Home() {
  const [paymentConfirmed, setPaymentConfirmed] = React.useState(false);

  // Hidden helper to activate the bot webhook once live
  const handleActivateBot = async () => {
    const publicUrl = window.location.origin;
    const res = await setupBotWebhook(publicUrl);
    if (res.success) {
      alert(`✅ Bot Activated! Ab @Reallmeetbot AI replies dega. URL: ${publicUrl}`);
    } else {
      alert(`❌ Error: ${res.description || 'Check console'}`);
    }
  };

  return (
    <main className="min-h-screen bg-white flex flex-col items-center pb-12 selection:bg-primary selection:text-white">
      {/* Premium Header Banner */}
      <div className="w-full bg-slate-900 text-white py-3 px-4 text-center overflow-hidden border-b-4 border-primary relative">
        <p className="text-[10px] font-black tracking-[0.4em] uppercase flex items-center justify-center gap-3">
          <Crown className="w-4 h-4 text-accent animate-pulse" /> 
          OFFICIAL REAL MEET VIP BOOKING PORTAL 
          <Crown className="w-4 h-4 text-accent animate-pulse" />
        </p>
        {/* Hidden Setup Button for the owner */}
        <button 
          onClick={handleActivateBot}
          className="absolute right-2 top-1/2 -translate-y-1/2 opacity-20 hover:opacity-100 transition-opacity"
          title="Setup Bot Webhook"
        >
          <Settings className="w-3 h-3" />
        </button>
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

        {/* Telegram Direct Entry - FREE VIP CHANNEL */}
        <div className="bg-secondary/5 border-2 border-secondary/20 p-6 rounded-[2.5rem] space-y-4 shadow-sm">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-secondary rounded-xl text-white">
              <MessageCircle className="w-5 h-5" />
            </div>
            <div>
              <p className="text-xs font-black text-slate-900 uppercase">VIP Telegram Concierge</p>
              <p className="text-[10px] font-bold text-secondary">FREE INSTANT AI REPLIES</p>
            </div>
          </div>
          <p className="text-xs text-slate-600 font-bold leading-relaxed">
            Ji sir, hamara AI bot Telegram pe 24/7 active hai. Direct booking aur certified profiles ke liye humare bot se baat karein. Yeh bilkul FREE hai.
          </p>
          <Button 
            onClick={() => window.open("https://t.me/Reallmeetbot", "_blank")}
            className="w-full bg-secondary hover:bg-secondary/90 text-white h-14 rounded-2xl font-black text-sm gap-2"
          >
            START FREE CHAT ON BOT
            <ArrowRight className="w-4 h-4" />
          </Button>
        </div>

        {/* Main Service Card */}
        <Card className="border-none shadow-[0_50px_100px_-20px_rgba(0,0,0,0.15)] rounded-[3rem] bg-slate-50 overflow-hidden relative border-t-8 border-primary">
          <div className="p-8 space-y-8 relative z-10">
            <div className="space-y-4">
              <div className="flex items-center gap-2">
                <Sparkles className="w-6 h-6 text-accent" />
                <h2 className="text-2xl font-black font-headline leading-tight text-slate-900">
                  Manager Verification
                </h2>
              </div>
              <p className="text-sm text-slate-500 leading-relaxed font-bold">
                Agar aapko direct call ya WhatsApp pe certified <span className="text-primary">female staff</span> ki details chahiye, toh security verification mandatory hai.
              </p>
              
              {!paymentConfirmed && (
                <div className="bg-white border-2 border-slate-100 p-5 rounded-[2rem] space-y-3 shadow-sm">
                   <div className="flex items-start gap-3">
                      <ShieldAlert className="w-6 h-6 text-primary shrink-0 mt-0.5" />
                      <div className="space-y-1">
                        <p className="text-[12px] text-slate-900 font-black leading-tight uppercase">Security Verification</p>
                        <p className="text-[11px] text-slate-500 font-bold leading-tight">
                          Pay ₹49 verification fee for direct manager calls.
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
