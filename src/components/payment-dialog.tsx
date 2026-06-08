
"use client";

import React from "react";
import { CheckCircle2, ShieldCheck, CreditCard, Loader2, Lock, ArrowRight } from "lucide-react";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogDescription,
} from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { Progress } from "@/components/ui/progress";
import { notifyTelegram } from "@/app/actions/telegram";

interface PaymentDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  onSuccess: () => void;
}

export function PaymentDialog({ open, onOpenChange, onSuccess }: PaymentDialogProps) {
  const [step, setStep] = React.useState<"initial" | "processing" | "success">("initial");
  const [progress, setProgress] = React.useState(0);

  const startRazorpayPayment = async () => {
    // Notify admin via Telegram that someone started payment
    await notifyTelegram("<b>New Payment Attempt:</b> A user clicked 'Pay Now' for the ₹49 fee.");

    setStep("processing");
    setProgress(0);
    
    // Simulating Razorpay SDK Loading & Verification
    const interval = setInterval(() => {
      setProgress((prev) => {
        if (prev >= 100) {
          clearInterval(interval);
          setStep("success");
          notifyTelegram("<b>Payment Successful:</b> User has successfully paid ₹49.");
          return 100;
        }
        return prev + 12;
      });
    }, 200);
  };

  const handleFinish = () => {
    onSuccess();
    setTimeout(() => {
      setStep("initial");
      setProgress(0);
    }, 500);
  };

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-md bg-white border-none rounded-[2rem] overflow-hidden shadow-2xl p-0">
        <div className="bg-primary p-8 text-center text-white space-y-2">
          <DialogTitle className="font-headline text-2xl font-black">
            {step === "success" ? "Verification Success" : "Security Verification"}
          </DialogTitle>
          <p className="text-xs font-bold opacity-80 uppercase tracking-widest">
            Real Meet Luxury Spa Portal
          </p>
        </div>

        <div className="p-8 space-y-6 flex flex-col items-center">
          {step === "initial" && (
            <>
              <div className="w-20 h-20 bg-primary/5 rounded-3xl flex items-center justify-center text-primary relative">
                <Lock className="w-10 h-10" />
                <div className="absolute -bottom-2 -right-2 bg-accent p-1.5 rounded-full border-4 border-white">
                  <ShieldCheck className="w-4 h-4 text-accent-foreground" />
                </div>
              </div>
              
              <div className="text-center space-y-1">
                <span className="text-5xl font-black italic text-slate-900 tracking-tighter">₹49</span>
                <p className="text-[10px] font-black text-slate-400 uppercase tracking-[0.2em]">Mandatory Security Fee</p>
              </div>

              <div className="w-full space-y-3">
                <Button 
                  onClick={startRazorpayPayment} 
                  className="w-full bg-primary hover:bg-primary/90 text-white font-black h-16 rounded-2xl shadow-xl shadow-primary/20 text-lg flex items-center justify-between px-6"
                >
                  <span>PAY VIA RAZORPAY</span>
                  <ArrowRight className="w-5 h-5" />
                </Button>
                <div className="flex items-center justify-center gap-4 text-[10px] font-bold text-slate-400">
                  <span className="flex items-center gap-1"><CreditCard className="w-3 h-3" /> UPI</span>
                  <span className="flex items-center gap-1"><CreditCard className="w-3 h-3" /> CARDS</span>
                  <span className="flex items-center gap-1"><CreditCard className="w-3 h-3" /> WALLETS</span>
                </div>
              </div>

              <p className="text-[10px] text-slate-400 text-center font-medium leading-relaxed max-w-[240px]">
                This fee prevents spam and ensures only serious users access our staff coordinates.
              </p>
            </>
          )}

          {step === "processing" && (
            <div className="w-full py-10 space-y-6 text-center">
              <div className="relative inline-flex">
                <Loader2 className="w-16 h-16 animate-spin text-primary" />
                <div className="absolute inset-0 flex items-center justify-center">
                  <div className="w-2 h-2 bg-primary rounded-full" />
                </div>
              </div>
              <div className="space-y-3">
                <p className="text-sm font-black text-primary animate-pulse">VERIFYING TRANSACTION...</p>
                <Progress value={progress} className="h-2 bg-slate-100 rounded-full" />
                <p className="text-[10px] text-slate-400 font-bold">PLEASE DO NOT REFRESH THIS PAGE</p>
              </div>
            </div>
          )}

          {step === "success" && (
            <>
              <div className="w-24 h-24 bg-secondary/10 rounded-full flex items-center justify-center text-secondary animate-in zoom-in duration-500">
                <CheckCircle2 className="w-14 h-14" />
              </div>
              <div className="text-center space-y-2">
                <p className="text-xl font-black text-secondary uppercase">Access Unlocked!</p>
                <p className="text-xs text-slate-500 font-bold">Redirecting you to the VIP Booking Console...</p>
              </div>
              <Button 
                onClick={handleFinish} 
                className="w-full bg-secondary hover:bg-secondary/90 text-white font-black h-16 rounded-2xl shadow-xl shadow-secondary/20"
              >
                START BOOKING NOW
              </Button>
            </>
          )}
        </div>
      </DialogContent>
    </Dialog>
  );
}
