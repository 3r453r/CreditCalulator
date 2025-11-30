# Grace Period Feature - Usage Guide

## Issue Analysis

Based on your server response, the problem is that `GracePeriodMonths` is set to `0` instead of `4`. This means the grace period feature is not being activated.

### Your Current Setup (INCORRECT)
```json
{
  "parameters": {
    "CreditStartDate": "2025-09-01",
    "CreditEndDate": "2026-05-30",      // Only 9 months, not 12!
    "GracePeriodMonths": 0,              // Should be 4!
    ...
  }
}
```

**Problems:**
1. `GracePeriodMonths` is `0` - the grace period feature is not enabled
2. Credit duration is only 9 months (Sep 2025 - May 2026), not 12 months as intended
3. With 0 grace period, principal payments start immediately from the first payment

### Correct Setup for 12-Month Credit with 4-Month Grace Period

```json
{
  "parameters": {
    "NetValue": 100000,
    "CreditStartDate": "2025-09-01",
    "CreditEndDate": "2026-08-31",      // Full 12 months
    "GracePeriodMonths": 4,              // 4 payments will be interest-only
    "PaymentFrequency": "Monthly",
    "PaymentDay": "LastOfMonth",
    ...
  },
  "rates": [
    {
      "DateFrom": "2025-09-01",
      "DateTo": "2025-09-09",
      "Rate": 3
    },
    {
      "DateFrom": "2025-09-10",
      "DateTo": "2025-11-29",
      "Rate": 4
    },
    {
      "DateFrom": "2025-11-30",
      "DateTo": "2026-08-31",          // Must cover entire credit period!
      "Rate": 5
    }
  ]
}
```

## How Grace Period Works

### Payment Count vs Calendar Months
Grace period is based on **number of payments**, not calendar months:
- `GracePeriodMonths: 4` means the **first 4 payments** are interest-only
- This is intentional to handle credits starting mid-month correctly

### Example with 12-Month Credit + 4-Month Grace:
- **Total duration**: 12 months (Sep 2025 - Aug 2026)
- **Payment frequency**: Monthly (Last of month)
- **Total payments**: 12 payments
- **Grace period**: 4 payments (first 4 are interest-only)
- **Principal repayment**: Last 8 payments

**Expected Schedule:**
1. **Payment 1 (Oct 31, 2025)**: Interest only, NO principal (Grace period)
2. **Payment 2 (Nov 30, 2025)**: Interest only, NO principal (Grace period)
3. **Payment 3 (Dec 31, 2025)**: Interest only, NO principal (Grace period)
4. **Payment 4 (Jan 31, 2026)**: Interest only, NO principal (Grace period)
5. **Payment 5 (Feb 28, 2026)**: Interest + Principal = 100,000 / 8 = 12,500
6. **Payment 6 (Mar 31, 2026)**: Interest + Principal = 12,500
7. ... (continues for remaining payments)
8. **Payment 12 (Aug 31, 2026)**: Final payment with remaining principal

## Common Mistakes

### ❌ Mistake 1: Not Setting GracePeriodMonths
If you don't enter a value in the "Karencja (miesiące)" field, it defaults to 0.
**Solution**: Enter `4` in the grace period field in the UI.

### ❌ Mistake 2: Credit Duration Too Short
If your credit is only 9 months with 4 months grace, you only have 5 months to repay principal.
**Solution**: For 12-month credit, set end date to 12 months after start date.

### ❌ Mistake 3: Rates Table Doesn't Cover Full Period
The rates table must cover the **entire credit period**, including the grace period, because interest still accrues during grace.
**Solution**: Last rate period's `DateTo` must match or exceed `CreditEndDate`.

## Rates Table and Grace Period

**IMPORTANT**: The rates table is **independent** of grace period:
- ✅ Rates table should cover the entire credit period from start to end
- ✅ Interest accrues normally during grace period using the rates table
- ✅ Grace period only controls whether principal is paid, not how interest is calculated

You do **NOT** need to adjust the rates table when changing grace period!

## How to Fix Your Issue

### Step 1: Set Credit Duration to 12 Months
Change your end date from `2026-05-30` to `2026-08-31` (or one year from your start date).

### Step 2: Set Grace Period to 4
In the UI, enter `4` in the "Karencja (miesiące)" field.

### Step 3: Extend Rates Table
Make sure your last rate period ends on or after `2026-08-31`.

### Step 4: Calculate
Click "Oblicz harmonogram" and verify:
- First 4 payments should have `PrincipalPayment: 0` (interest only)
- Payments 5-12 should have `PrincipalPayment: 12500` (principal + interest)
- `RemainingPrincipal` should stay at 100,000 for first 4 payments

## Verification Checklist

After fixing, verify your request contains:
- ✅ `"GracePeriodMonths": 4` (not 0)
- ✅ `"CreditEndDate": "2026-08-31"` (12 months from start)
- ✅ Last rate period ends on or after credit end date
- ✅ First 4 schedule items have `"PrincipalPayment": 0`
- ✅ Remaining schedule items have principal payments

## Test File

A correct example request is provided in `test-grace-period-example.json`. You can:
1. Import this file using the "Import" feature
2. Click "Oblicz harmonogram" to see the correct grace period behavior
3. Verify that the first 4 payments are interest-only
