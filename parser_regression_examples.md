# Parser Regression Examples

Purpose: keep real customer-order examples that must continue to parse after parser changes.

Rule for future parser edits: after changing `app0518_smart_local.py`, rerun these examples and verify the expected critical SKUs still appear. Fuzzy guesses should stay in review unless explicitly expected as exact matches.

## Example 1 - Missing C Prefix For YN/MG Variants

Expected:
- `YNGBP15K-BB` parses as `C-YNGBP15K-BB`
- existing `B-`, `P-`, `D-`, `G-`, `C-` prefixes stay unchanged

```text
C-GBP15K-DM	1.000
C-GBP15K-MMI	1.000
C-GBP15K-SK	1.000
C-GBPX25-BB	1.000
C-GBPX25-CM	1.000
C-GBPX25-MM	1.000
C-GBPX25-OD	1.000
C-GBPX25-PB	1.000
C-GBPX25-POT	1.000
C-GBPX25-SD	1.000
C-GBPX25-SJ	1.000
C-GBPX25-SMP	1.000
YNGBP15K-BB	1.000
YNGBP15K-BC	1.000
YNGBP15K-BW	1.000
YNGBP15K-CC	1.000
YNGBP15K-SG	1.000
YNGBP15K-SM	1.000
YNGBP15K-TRB	1.000
YNGBP15K-WGI	1.000
```

Critical outputs:

```text
C-YNGBP15K-BB
C-YNGBP15K-BC
C-YNGBP15K-BW
C-YNGBP15K-CC
C-YNGBP15K-SG
C-YNGBP15K-SM
C-YNGBP15K-TRB
C-YNGBP15K-WGI
```

## Example 2 - IJOY XP And XCube Ice Pod

Expected:
- `IJOY KIT MG_XP100000` maps to `IJOY KIT MG_XP100000`
- `IJOY_XP50000` maps to `IJOY_XP50000`
- `XCube ICe POD` maps to `OFF STAMP X-CUBE POD (ICE)_25000`
- `Pomegranate Blast` parses as `C-XC25IP-PB`
- `Pink & Blue` parses as `C-MGIJXP100KIT-P&B`

```text
IJOY KIT MG_XP100000
- Blueberry Blackberry
- Blue Cotton Candy
- Blue Razz Ice
- Blueberry Watermelon
- Cool Mint
- Grape Ice
- Mixed Berries
- Miami Mint
- Pink & Blue
- Rose Donut
- Strawberry Ice
- Watermelon Ice

IJOY_XP50000
- Blue Razz Ice
- Blueberry Watermelon
- Miami Mint
- Triple Berry

Off Stamp X-Cube35000 Crystal Pod(150 bars / carton)
Black Mint
Blackberry Blueberry
Blueberry Sour Raspberry
Cool Mint
Gummy Bear
Mountain Berry
Strawberry Cupcake
Watermelon Bubble Gum
XCube ICe POD
Pomegranate Blast
```

Critical outputs:

```text
C-MGIJXP100KIT-P&B
C-IJXP50-TB
C-XC25CSPOD-BM
C-XC25CSPOD-WBG
C-XC25IP-PB
```

## Example 6 - Lost Mary MT35 Percent Header, Case Units, NERA, XCube Mixed Quantities

Observed issue:
- `Lostmary MT35 0%` is a separate catalog family: `LOST MARY (0%)_MT35000`, SKU prefix `C-LMMT350-`; do not treat it as regular MT35.
- Lines ending in `1 case` should parse quantity 1 and remove the unit from the flavor text.
- `Lostmary MT35k` without `0%` should map to regular/Lemonade MT35 candidates, not the 0% family.
- Mixed quantity blocks should follow current rule: if some lines have quantities and some do not, no-quantity flavor lines are review/0, not auto-added as quantity 1.
- Exact `+` flavors that exist in the selected family should auto-add; unavailable or cross-family fuzzy guesses must require confirmation.

```text
Lostmary MT35 0%
Blue Razz Ice  1 case 
Baja Splash  1 case
Miami Mint  1 case 
Watermelon+  1 case 
Winter Mint 1  case 

Lostmary MT35k
Blue Razz Lemonade 1 case 
Strawberry Dragon Lemonade  1 case 
Tropical Lemonade  1 case 
Miami Mint 1
Strawberry+ 1
Watermelon+ 1
Black Mint+  1
Blue Razz Cotton Carnival 1
Blue Razz Ice  1
Classic Tobacco 1
Mountain Berry 1

Lostmary nera 70k kit
Blue Razz Ice  1 case 
Baja Splash & Hawaiian Punch 1 case 
Cool Mint & Polar Mint 1 case 
Dragon Strawnana 1 case 
Juicy Peach Ice & Pineapple Coconut 1 case 
Miami Mint 1 case 
Miami Mint & Winter Mint 1 case 
Strawberry Ice & Strawberry  Watermelon 1 case 

Lostmary Nera70k pod 
Blackberry Blueberry 1 case 
Baja Splash 1 case 
Clear 1 case 
Dragon strawnana 1 case 
Golden Berry 1 case
Pomegranate Blast 1 case 
Pink & Blue 1 case 
Pineapple Coconut 1 case 
Pineapple Passion Orange 1 case 
Rocket Freeze 1 case 
Scary Berry 1 case

Xcube crystal pod 35k
Blueberry Apple Cobbler  1 case 
Black Mint 1 case 
Florida Orange 1 case 
Mountain Berry 1 case
Peach Berry 1 case 
Sour Apple Ice 1 case 
Vanilla 1 case 
Watermelon Ice 1
Mexican Mango 1
White Gummy 1
Cool Mint 1
Clear 1
Watermelon Bubble Gum 1
Pineapple Coconut 1 
Strawberry Cupcake 1
Gummy Bear 1
Grape 1
Blueberry Sour Raspberry
Rocket Freeze
```

Critical exact outputs to preserve after fixing:

```text
C-LMMT350-BRI
C-LMMT350-BS
C-LMMT350-MM
C-LMMT350-W+
C-LMMT350-WM
C-LMMT35LE-BRL
C-LMMT35LE-SDL
C-LMMT35LE-TL
C-LMNFK70-BRI
C-LMNFK70-BSHP
C-LMNFK70-CMPM
C-LMNFK70-DS
C-LMNFK70-JPPC
C-LMNFK70-MM
C-LMNFK70-MMWM
C-LMNFK70-SISW
C-LMNFP70-BB
C-LMNFP70-BS
C-LMNFP70-CLEA
C-LMNFP70-DS
C-LMNFP70-GB
C-LMNFP70-PB
C-LMNFP70-PBLU
C-LMNFP70-PC
C-LMNFP70-PPO
C-LMNFP70-RF
C-LMNFP70-SB
C-XC25CSPOD-BAC
C-XC25CSPOD-BM
C-XC25CSPOD-FO
C-XC25CSPOD-MB
C-XC25CSPOD-PB
C-XC25CSPOD-SAI
C-XC25CSPOD-VANI
C-XC25CSPOD-WI
C-XC25CSPOD-MMG
C-XC25CSPOD-WG
C-XC25CSPOD-CM
C-XC25CSPOD-CLEA
C-XC25CSPOD-WBG
C-XC25CSPOD-PC
C-XC25CSPOD-SC
C-XC25CSPOD-GB
C-XC25CSPOD-GRAP
```

Expected review candidates:

```text
Blueberry Sour Raspberry
Rocket Freeze
```

## Example 3 - XCube Crystal Kit Without Off Stamp Prefix

Expected:
- `X-Cube 35000 CRYSTAL Kit` maps to `OFF STAMP X-CUBE KIT (CRYSTAL)_35000`
- Each flavor defaults to quantity 1 when no quantities are listed

```text
X-Cube 35000 CRYSTAL Kit
Blackberry Blueberry
Blue Razz Ice
Mountain Berry
Winter Mint
```

Critical outputs:

```text
C-XC25CSKIT-BB
C-XC25CSKIT-BRI
C-XC25CSKIT-MB
C-XC25CSKIT-WM
```

## Example 4 - MT35K With Lost Mary Prefix

Expected:
- `lost mary MT35k` maps to `LOST MARY_MT35000`
- Exact MT35 flavors parse automatically
- Typos remain review candidates unless exact catalog normalization is available

```text
lost mary  MT35k
Baja splash 2
Blackberry blueberry 1
Blue RAZ ice 1
Black Raz lemon 2
Black mint 2
Golden Bery 1
Kiwi passion fruit 1
Mountains berry 2
Rocket freeze 1
Scary berry 1
Sunny orange 2
White Gami 1
Winter mint 2
Miami mint 2
Red pixy 2
White pixy 1
```

Critical exact outputs:

```text
C-LMMT35-BS
C-LMMT35-KPF
C-LMMT35-RF
C-LMMT35-SO
C-LMMT35-WM
C-LMMT35-MM
C-LMMT35-RP
C-LMMT35-WP
```

Expected review candidates:

```text
Blackberry blueberry
Blue RAZ ice
Black Raz lemon
Black mint
Golden Bery
Mountains berry
White Gami
```

## Example 5 - Classic XCube Pod Should Search Classic, Sweet, And Ice

Expected:
- `Classic-XCube Pod` searches Classic, Sweet, and Ice Pod families
- `Pomegranate Blast` resolves to Ice Pod SKU `C-XC25IP-PB`

```text
Classic-XCube Pod (150
Pomegranate Blast
```

Critical outputs:

```text
C-XC25IP-PB
```
