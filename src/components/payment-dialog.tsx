"use client";

import React from "react";
import { CheckCircle2, ShieldCheck, CreditCard, Loader2, Lock } from "lucide-react";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogDescription,
} from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { Progress } from "@/components/ui/progress";

interface PaymentDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  onSuccess: () => void;
}

export function PaymentDialog({ open, onOpenChange, onSuccess }: PaymentDialogProps) {
  const [step, setStep] = React.useState<"initial" | "processing" | "success">("initial");
  const [progress, setProgress] = React.useState(0);

  const startPayment = () => {
    setStep("processing");
    setProgress(0);
    const interval = setInterval(() => {
      setProgress((prev) => {
        if (prev >= 100) {
          clearInterval(interval);
          setStep("success");
          return 100;
        }
        return prev + 8;
      });
    }, 150);
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
      <DialogContent className="sm:max-w-md bg-white border-none rounded-3xl overflow-hidden shadow-2xl">
        <DialogHeader className="p-6 pb-0">
          <DialogTitle className="font-headline text-2xl text-primary text-center">
            {step === "success" ? "Verification Complete" : "Join Our VIP Portal"}
          </DialogTitle>
          <DialogDescription className="text-center text-sm font-medium">
            {step === "success" 
              ? "Your access to verified female staff is now active." 
              : "Pay a nominal ₹49 registration fee to access verified staff profiles and private call slots."}
          </DialogDescription>
        </DialogHeader>

        <div className="flex flex-col items-center justify-center p-8 space-y-6">
          {step === "initial" && (
            <>
              <div className="relative">
                <div className="w-24 h-24 bg-primary/5 rounded-full flex items-center justify-center text-primary">
                  <Lock className="w-12 h-12" />
                </div>
                <div className="absolute -bottom-1 -right-1 bg-accent p-1 rounded-full text-accent-foreground border-4 border-white">
                  <ShieldCheck className="w-5 h-5" />
                </div>
              </div>
              <div className="text-center">
                <span className="text-5xl font-black font-headline text-primary">₹49</span>
                <p className="text-[10px] font-bold text-muted-foreground uppercase tracking-widest mt-2">One-time Verification Fee</p>
              </div>
              <Button onClick={startPayment} className="w-full bg-primary hover:bg-primary/90 text-primary-foreground font-bold h-14 rounded-2xl shadow-lg text-lg">
                Pay Securely Now
              </Button>
              <p className="text-[10px] text-muted-foreground flex items-center gap-1 font-medium">
                <ShieldCheck className="w-3 h-3 text-secondary" />
                SSL Encrypted Payment Gateway
              </p>
            </>
          )}

          {step === "processing" && (
            <>
              <div className="w-20 h-20 relative flex items-center justify-center">
                <Loader2 className="w-14 h-14 animate-spin text-primary" />
              </div>
              <div className="w-full space-y-3 text-center">
                <p className="text-sm font-bold text-primary animate-pulse">Verifying Transaction...</p>
                <Progress value={progress} className="h-2 bg-muted rounded-full" />
              </div>
            </>
          )}

          {step === "success" && (
            <>
              <div className="w-24 h-24 bg-secondary/10 rounded-full flex items-center justify-center text-secondary animate-in zoom-in duration-500">
                <CheckCircle2 className="w-14 h-14" />
              </div>
              <div className="text-center space-y-2">
                <p className="text-xl font-bold text-secondary">Verified Successfully!</p>
                <p className="text-xs text-muted-foreground font-medium">Redirecting you to the spa coordinator...</p>
              </div>
              <Button onClick={handleFinish} className="w-full bg-secondary hover:bg-secondary/90 text-secondary-foreground font-bold h-14 rounded-2xl shadow-lg">
                Enter VIP Booking
              </Button>
            </>
          )}
        </div>
      </DialogContent>
    </Dialog>
  );
}