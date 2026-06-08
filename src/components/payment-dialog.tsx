"use client";

import React from "react";
import { CheckCircle2, ShieldCheck, CreditCard, Loader2 } from "lucide-react";
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
        return prev + 5;
      });
    }, 100);
  };

  const handleFinish = () => {
    onSuccess();
    // Reset state after closure
    setTimeout(() => {
      setStep("initial");
      setProgress(0);
    }, 500);
  };

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-md bg-card border-border">
        <DialogHeader>
          <DialogTitle className="font-headline text-2xl">
            {step === "success" ? "Payment Confirmed" : "Unlock Professional Booking"}
          </DialogTitle>
          <DialogDescription>
            {step === "success" 
              ? "Your session is now ready to be scheduled." 
              : "Pay a one-time fee of ₹49 to access professional call slots."}
          </DialogDescription>
        </DialogHeader>

        <div className="flex flex-col items-center justify-center py-6 space-y-6">
          {step === "initial" && (
            <>
              <div className="w-20 h-20 bg-primary/10 rounded-full flex items-center justify-center text-primary">
                <CreditCard className="w-10 h-10" />
              </div>
              <div className="text-center">
                <span className="text-4xl font-bold font-headline">₹49</span>
                <p className="text-sm text-muted-foreground mt-1">One-time gateway fee</p>
              </div>
              <Button onClick={startPayment} className="w-full bg-primary hover:bg-primary/90 text-primary-foreground font-semibold h-12">
                Pay Now
              </Button>
              <div className="flex items-center gap-2 text-xs text-muted-foreground">
                <ShieldCheck className="w-4 h-4" />
                Secure 256-bit SSL Transaction
              </div>
            </>
          )}

          {step === "processing" && (
            <>
              <div className="w-20 h-20 relative flex items-center justify-center">
                <Loader2 className="w-12 h-12 animate-spin text-primary" />
              </div>
              <div className="w-full space-y-2 text-center">
                <p className="text-sm font-medium">Processing secure payment...</p>
                <Progress value={progress} className="h-2" />
              </div>
            </>
          )}

          {step === "success" && (
            <>
              <div className="w-20 h-20 bg-green-500/10 rounded-full flex items-center justify-center text-green-500 animate-in zoom-in duration-300">
                <CheckCircle2 className="w-12 h-12" />
              </div>
              <div className="text-center space-y-2">
                <p className="text-lg font-semibold">Transaction Successful!</p>
                <p className="text-sm text-muted-foreground">Redirecting you to the booking system...</p>
              </div>
              <Button onClick={handleFinish} className="w-full bg-primary hover:bg-primary/90 text-primary-foreground font-semibold h-12">
                Go to Booking
              </Button>
            </>
          )}
        </div>
      </DialogContent>
    </Dialog>
  );
}