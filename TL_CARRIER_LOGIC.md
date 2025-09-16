# TL Carrier Logic Implementation

## Overview
The BVC Automator now includes special handling for TL (Truck Load) carriers as specified in the "UPDATED LTL + TL LIST.xlsx" file.

## TL Carriers
The following carriers receive special "copy-paste and zero out" treatment:
- **LANDSTAR RANGER INC**
- **SMARTWAY TRANSPORTATION INC**

## Business Logic
When the system encounters any of the TL carriers (either in Selected Carrier or Least Cost Carrier columns):

1. **Copy Operation**: Selected carrier data is copied to all Least Cost columns:
   - Selected Carrier → Least Cost Carrier
   - Selected Service Type → Least Cost Service Type
   - Selected Transit Days → Least Cost Transit Days
   - Selected Freight Cost → Least Cost Freight Cost
   - Selected Accessorial Cost → Least Cost Accessorial Cost
   - Selected Total Cost → Least Cost Total Cost

2. **Zero Out**: Potential Savings is set to $0.00

## Implementation
- **Basic Reports**: Applied in `ModernTMSProcessor._apply_business_logic_enhanced()` as Rule 4
- **Detailed Reports**: Applied in `TMSDetailedDataProcessor._apply_business_logic_detailed()` as Rule 4
- **Logging**: Actions are logged with count of affected rows

## Purpose
This logic ensures that TL carriers (which operate under different business models than LTL carriers) are processed correctly, with savings opportunities zeroed out as they follow different pricing structures.

## LTL vs TL
- **LTL LIST**: 73 carriers that follow standard business logic (normal processing)
- **TL LIST**: 2 carriers that require special copy-paste-zero treatment

The tool's primary purpose is to apply logic to LTL carriers only, with special handling for the few TL carriers that may appear in the data.