"use client";

import React from "react";
import { Calendar as CalendarIcon, Clock, CheckCircle2 } from "lucide-react";
import { Card, CardHeader, CardTitle, CardContent } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";

interface Slot {
  id: string;
  time: string;
  display: string;
}

interface BookingSlotsProps {
  slots: Slot[];
  onSelect: (slot: Slot) => void;
}

export function BookingSlots({ slots, onSelect }: BookingSlotsProps) {
  const [selectedId, setSelectedId] = React.useState<string | null>(null);

  return (
    <Card className="border-border bg-card/30 backdrop-blur-md">
      <CardHeader className="pb-3">
        <CardTitle className="text-lg font-headline flex items-center gap-2">
          <CalendarIcon className="w-5 h-5 text-primary" />
          Available Call Slots
        </CardTitle>
      </CardHeader>
      <CardContent>
        <div className="grid grid-cols-1 gap-2">
          {slots.length > 0 ? (
            slots.map((slot) => (
              <Button
                key={slot.id}
                variant={selectedId === slot.id ? "default" : "outline"}
                className={cn(
                  "h-auto py-3 px-4 justify-between font-normal transition-all duration-200",
                  selectedId === slot.id ? "border-primary" : "border-border"
                )}
                onClick={() => {
                  setSelectedId(slot.id);
                  onSelect(slot);
                }}
              >
                <div className="flex items-center gap-3">
                  <Clock className="w-4 h-4 text-primary" />
                  <span className="text-sm">{slot.display}</span>
                </div>
                {selectedId === slot.id && <CheckCircle2 className="w-4 h-4 text-primary-foreground" />}
              </Button>
            ))
          ) : (
            <p className="text-sm text-muted-foreground py-4 text-center">
              Fetching latest availability...
            </p>
          )}
        </div>
        
        {selectedId && (
          <div className="mt-4 animate-in fade-in slide-in-from-top-2 duration-300">
            <Button className="w-full bg-primary hover:bg-primary/90 text-primary-foreground font-semibold h-10">
              Confirm Booking
            </Button>
          </div>
        )}
      </CardContent>
    </Card>
  );
}