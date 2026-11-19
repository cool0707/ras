/* Ported from xlflow (https://github.com/harumiWeb/xlflow), MIT License,
   Copyright (c) 2026 harumiWeb. See THIRD_PARTY_NOTICES.md. */
import type { PropertyDescriptor } from "../../../src/userForm/protocol";
import { numberRoundTrips } from "../numbers";

export function validNumericInput(text: string, descriptor: PropertyDescriptor): boolean {
  const number = Number(text);
  return (
    text.trim() !== "" &&
    Number.isFinite(number) &&
    numberRoundTrips(text, number) &&
    (descriptor.valueType !== "integer" || Number.isSafeInteger(number)) &&
    (descriptor.minimum === undefined ||
      (descriptor.exclusiveMinimum ? number > descriptor.minimum : number >= descriptor.minimum)) &&
    (descriptor.maximum === undefined || number <= descriptor.maximum)
  );
}
