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
