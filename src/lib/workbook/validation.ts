import type { CellData, DataValidation } from "@/lib/workbook/types";

export function validateCellValue(
  input: string,
  validation?: DataValidation,
): { valid: boolean; message?: string } {
  if (!validation) return { valid: true };
  const trimmed = input.trim();
  if (!trimmed && validation.allowBlank) return { valid: true };
  if (!trimmed) {
    return {
      valid: false,
      message: validation.errorMessage ?? "This cell cannot be blank.",
    };
  }

  switch (validation.kind) {
    case "list": {
      const allowed = validation.values ?? [];
      if (!allowed.includes(trimmed)) {
        return {
          valid: false,
          message:
            validation.errorMessage ??
            `Value must be one of: ${allowed.join(", ")}`,
        };
      }
      return { valid: true };
    }
    case "whole": {
      const n = Number(trimmed);
      if (!Number.isInteger(n)) {
        return { valid: false, message: validation.errorMessage ?? "Whole number required." };
      }
      if (validation.min !== undefined && n < validation.min) {
        return { valid: false, message: validation.errorMessage ?? `Minimum is ${validation.min}.` };
      }
      if (validation.max !== undefined && n > validation.max) {
        return { valid: false, message: validation.errorMessage ?? `Maximum is ${validation.max}.` };
      }
      return { valid: true };
    }
    case "decimal": {
      const n = Number(trimmed);
      if (Number.isNaN(n)) {
        return { valid: false, message: validation.errorMessage ?? "Decimal number required." };
      }
      if (validation.min !== undefined && n < validation.min) {
        return { valid: false, message: validation.errorMessage ?? `Minimum is ${validation.min}.` };
      }
      if (validation.max !== undefined && n > validation.max) {
        return { valid: false, message: validation.errorMessage ?? `Maximum is ${validation.max}.` };
      }
      return { valid: true };
    }
    case "textLength": {
      const len = trimmed.length;
      if (validation.min !== undefined && len < validation.min) {
        return { valid: false, message: validation.errorMessage ?? `Minimum length is ${validation.min}.` };
      }
      if (validation.max !== undefined && len > validation.max) {
        return { valid: false, message: validation.errorMessage ?? `Maximum length is ${validation.max}.` };
      }
      return { valid: true };
    }
    case "date":
      if (Number.isNaN(Date.parse(trimmed))) {
        return { valid: false, message: validation.errorMessage ?? "Invalid date." };
      }
      return { valid: true };
    default:
      return { valid: true };
  }
}

export function attachValidation(
  data: CellData,
  validation: DataValidation,
): CellData {
  return { ...data, validation };
}