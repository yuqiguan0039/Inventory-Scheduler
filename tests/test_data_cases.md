# Consolidated Order Parser Test Data

> Canonical regression fixture generated from `fixtures/order_parser/测试数据模板.xlsx`.
> Keep each original input block unchanged when adding a new failure case.
> For text cases, the `Expected SKU + Quantity` block is the machine-checkable assertion; lines without a catalog SKU are notes/review expectations.

## text_001 [text]
Prompt: 请总结下面这段文字
Image path: (none)
Difficulty: 简单
Tags: 总结；中文
Acceptance: 提到主要观点；不编造信息
Notes: 复制本行继续填写

### Input
```text
在这里粘贴原始文字
```

### Expected SKU + Quantity
```text
应包含核心结论
```

## image_text_001 [image_text]
Prompt: 请识别图片中的日期和总金额
Image path: images/example_001.png
Difficulty: 中等
Tags: 图片；OCR
Acceptance: 字段完整；识别准确；不确定时说明
Notes: 图片路径建议使用相对路径

### Input
```text
可选的补充说明
```

### Expected SKU + Quantity
```text
日期：2026-08-01；总金额：128.50
```

## Shirley001 [text]
Prompt: 请将这段brand,flavor,quantity 总结成 sku +quantity 模式
Image path: (none)
Difficulty: 困难
Tags: SKU匹配；组合口味；&分隔符；Kit/Pod；拼写错误；需人工确认
Acceptance: Kit 应输出 7 个 SKU，Pod 应输出 9 个已确认 SKU；每行默认数量为 1；“&”连接的整行是一个组合口味，不得拆分；Kit 与 Pod 不得混用 SKU 家族；Lemn mint 必须标记为待确认，不得猜测或套用其他产品 SKU。
Notes: 重点回归测试：Baja Splash & Hawaiian Punch 是单个组合口味，对应 C-LMNFK70-BSHP，不是 Baja Splash 和 Hawaiian Punch 两个口味。Blue Razz Ice & Blackberry Blueberry、Juicy Peach Ice & Pineapple Coconut、Strawberry Ice & Strawberry Watermelon、Watermelon Ice & Watermelon Fcuking Fab 也应按同一规则处理。

### Input
```text
LOST MARY NERA FULLVIEW Kit 70000 24ml (100 bars / carton) 
Baja Splash & Hawaiian Punch
Blue Razz Ice & Blackberry Blueberry
Dragon Strawnana
Juicy Peach Ice & Pineapple Coconut
Pineapple Passion Orange
Strawberry Ice & Strawberry Watermelon
Watermelon Ice & Watermelon Fcuking Fab
LOST MARY NERA FULLVIEW pod (100 bars / carton)
Baja Splash
Blackberry Blueberry
Cool Mint
Dragon Strawnana
Hawaiian Punch
Lemn mint 
Pink & Blue
Pomegranate Blast
Strawberry Watermelon
Watermelon Fcuking Fab
```

### Expected SKU + Quantity
```text
C-LMNFK70-BSHP	1
C-LMNFK70-BRBB	1
C-LMNFK70-DS	1
C-LMNFK70-JPPC	1
C-LMNFK70-PPO	1
C-LMNFK70-SISW	1
C-LMNFK70-WIWF	1
C-LMNFP70-BS	1
C-LMNFP70-BB	1
C-LMNFP70-CM	1
C-LMNFP70-DS	1
C-LMNFP70-HP	1
C-LMNFP70-PBLU	1
C-LMNFP70-PB	1
C-LMNFP70-SW	1
C-LMNFP70-WFF	1
Lemn mint	NEEDS REVIEW (no matching NERA FULLVIEW Pod catalog flavor)
```

## Shirley002 [text]
Prompt: 请将这段brand,flavor,quantity 总结成 sku +quantity 模式
Image path: (none)
Difficulty: 中等
Tags: SKU 匹配；组合口味；World Cup 家族；人工审核
Acceptance: 仅精确匹配自动加入；Florida twist (world cup) 在对应 MT35 World Cup 家族中没有目录 SKU，必须人工确认。
Notes: (none)

### Input
```text
Lost Mary MT35K 
Strawberry dragon lemonade 
Baja splash
Miami mint (world cup)
Red pixy
Summer grape
Florida twist (world cup)

Lost Mary nera 70k kit
Rocket freeze 
Dragon strawnana 
Blue raz ice & Blackberry Blueberry 

Lost Mary nera 70k pod
Watermelon ice 
Miami mint 
Lemon pop
Scary berry 

Off stamp crystal pod 35k
Rasberry watermelon 
Tropical punch
Blue raz grape ice 
Bangin watermelon
```

### Expected SKU + Quantity
```text
C-LMMT35LE-SDL	1
C-LMMT35-BS	1
C-LMMT35WC-MM	1
C-LMMT35-RP	1
C-LMMT35-SG	1
C-LMNFK70-RF	1
C-LMNFK70-DS	1
C-LMNFK70-BRBB	1
C-LMNFP70-WI	1
C-LMNFP70-MM	1
C-LMNFP70-LP	1
C-LMNFP70-SB	1
C-XC25CSPOD-RW	1
C-XC25CSPOD-TP	1
C-XC25CSPOD-BRGI	1
C-XC25CSPOD-BW	1
```

## MixedSKUHeaders_20260909 [text]
Prompt: 请将这段 SKU、产品标题和口味数量统一转换成 C-SKU + quantity
Image path: (none)
Difficulty: 困难
Tags: SKU匹配；无前缀SKU；反斜杠转义；HTML空格实体；Xcube Crystal Pod；MT35K；回归
Acceptance: 15 条 GBPX25、6 条 GBP15K 和两个产品区块的口味必须全部识别；无前缀 SKU 自动补 `C-`；`&#x20;`、`SKU` 和行尾反斜杠视为格式噪声；总计 41 条，数量合计 51，Need Review 为 0。
Notes: 新增回归样本。原文同时包含无前缀 SKU、复制产生的反斜杠和 HTML 空格实体。`Xcube crystal pod` 应命中 `OFF STAMP X-CUBE POD (CRYSTAL)_35000`；`Lostmary MT35k` 应命中普通 `LOST MARY_MT35000` 家族。弯引号 `Bangin’` 和目录中的撇号形式可正常匹配。

### Input
```text
GBPX25-BBP\
GBPX25-BTF\
GBPX25-MM\
GBPX25-OD\
GBPX25-OFF\
GBPX25-OM\
GBPX25-PBL\
GBPX25-PJ\
GBPX25-POT\
GBPX25-RPL\
GBPX25-SC\
GBPX25-SMP\
GBPX25-SS\
GBPX25-SW\
GBPX25-WCSS\
&#x20;\
&#x20;\
SKU
GBP15K-CMT\
GBP15K-PS\
GBP15K-SABP\
GBP15K-SS\
GBP15K-TRB\
GBP15K-WI\
Xcube crystal pod 
Bangin’ Strawberry 1\
Bangin’ Watermelon 1\
Blackberry Blueberry 1\
Blue Razz Ice 1\
Florida Orange 1\
Maui Punch 1\
Pineapple Coconut 1\
Raspberry Watermelon 1\
Strawberry Cupcake 1\
Tropical Punch 1\
Watermelon Bubble Gum 1\
Watermelon Ice 1\
Lostmary MT35k
Blue Razz Ice 2\
Miami Mint 2\
Mountain Berry 1\
Pink Lemonade+ 1\
Strawmelon Peach 2\
Summer Grape 2\
Tigers Blood+ 3\
Winter Mint 5
```

### Expected SKU + Quantity
```text
C-GBPX25-BBP	1
C-GBPX25-BTF	1
C-GBPX25-MM	1
C-GBPX25-OD	1
C-GBPX25-OFF	1
C-GBPX25-OM	1
C-GBPX25-PBL	1
C-GBPX25-PJ	1
C-GBPX25-POT	1
C-GBPX25-RPL	1
C-GBPX25-SC	1
C-GBPX25-SMP	1
C-GBPX25-SS	1
C-GBPX25-SW	1
C-GBPX25-WCSS	1
C-GBP15K-CMT	1
C-GBP15K-PS	1
C-GBP15K-SABP	1
C-GBP15K-SS	1
C-GBP15K-TRB	1
C-GBP15K-WI	1
C-XC25CSPOD-BS	1
C-XC25CSPOD-BW	1
C-XC25CSPOD-BB	1
C-XC25CSPOD-BRI	1
C-XC25CSPOD-FO	1
C-XC25CSPOD-MP	1
C-XC25CSPOD-PC	1
C-XC25CSPOD-RW	1
C-XC25CSPOD-SC	1
C-XC25CSPOD-TP	1
C-XC25CSPOD-WBG	1
C-XC25CSPOD-WI	1
C-LMMT35-BRI	2
C-LMMT35-MM	2
C-LMMT35-MB	1
C-LMMT35-PL+	1
C-LMMT35-SP	2
C-LMMT35-SG	2
C-LMMT35-TB+	3
C-LMMT35-WM	5
```

## Shirley0804001 [text]
Prompt: 请将这段brand,flavor,quantity 总结成 sku +quantity 模式
Image path: (none)
Difficulty: 困难
Tags: SKU 匹配；Classic/Crystal；模糊口味；人工审核
Acceptance: 仅精确匹配自动加入；Smashin Grape、Twist Strawberry、Peach Medley、Blue Razz Ice Cream、Mexico Mango 与当前 Classic/Crystal 家族存在目录差异或模糊匹配，必须人工确认。
Notes: (none)

### Input
```text
Off-Stamp Classic X-Cube PODs

* Bangin Watermelon
* Bangin Peach
* Dragon Strawnana
* Blue Razz Ice
* Florida Sherbet
* Hawaiian Popsicle
* Miami Mint
* Pineapple Passion Orange
* Rocket Freeze
* Smashin Grape
* Twist Strawberry
* Peach Medley

Off-Stamp X-Cube Crystal PODs

* Bangin Strawberry
* Bangin Watermelon
* Blue Razz Ice Cream
* Blueberry Apple Cobbler
* Blueberry Sour Raspberry
* Cool Mint
* Florida Orange
* Grape
* Gummy Bear
* Maui Punch
* Mexico Mango
* Miami Mint 
* Rocket Freeze 
* Sour Apple Ice
* Tropical Guava
* Vanilla
* Watermelon Bubblegum
* Watermelon Ice
* Winter Mint
```

### Expected SKU + Quantity
```text
C-XC25CP-BW	1
C-XC25CP-BP	1
C-XC25CP-DS	1
C-XC25CP-BRI	1
C-XC25CP-FS	1
C-XC25CP-HP	1
C-XC25CP-MM	1
C-XC25CP-PPO	1
C-XC25CP-RF	1
C-XC25CSPOD-BS	1
C-XC25CSPOD-BW	1
C-XC25CSPOD-BAC	1
C-XC25CSPOD-BSR	1
C-XC25CSPOD-CM	1
C-XC25CSPOD-FO	1
C-XC25CSPOD-GRAP	1
C-XC25CSPOD-GB	1
C-XC25CSPOD-MP	1
C-XC25CSPOD-MM	1
C-XC25CSPOD-RF	1
C-XC25CSPOD-SAI	1
C-XC25CSPOD-TG	1
C-XC25CSPOD-VANI	1
C-XC25CSPOD-WBG	1
C-XC25CSPOD-WI	1
C-XC25CSPOD-WM	1
```

## Aven0807001 [text]
Prompt: 请将这段brand,flavor,quantity 总结成 sku +quantity 模式
Image path: (none)
Difficulty: (not set)
Tags: (none)
Acceptance: (not set)
Notes: (none)

### Input
```text
Lost Mary MT35k 
Blue Razz Lemonade 1 
Mint Lemonade 2
Strawberry Dragon Lemonade 1 
Tropical Lemonade 2
Blackberry Blueberry 1 
Black Mint+ 1
Black Razz Lemon 1
Clear 3
Classic Tobacco 1
Golden Berry 1
Kiwi Passion Fruit 1
Miami Mint 2 
Scary Berry+ 2 
Summer Grape 2 
Strawberr kiwi 1 
Sunny Orange 1 
Strawmelon Peach 1 
Toasted Banana 1
Tigers Blood+ 2
Watermelon+ 1 
White Gami＋ 2 
Winter Mint 2
```

### Expected SKU + Quantity
```text
C-LMMT35LE-BRL	1
C-LMMT35LE-ML	2
C-LMMT35LE-SDL	1
C-LMMT35LE-TL	2
C-LMMT35-BB+	1
C-LMMT35-BM+	1
C-LMMT35-BRL	1
C-LMMT35-CLEA	3
C-LMMT35-CT	1
C-LMMT35-GB	1
C-LMMT35-KPF	1
C-LMMT35-MM	2
C-LMMT35-SB+	2
C-LMMT35-SG	2
C-LMMT35-SK+	1
C-LMMT35-SO	1
C-LMMT35-SP	1
C-LMMT35-TB	1
C-LMMT35-TB+	2
C-LMMT35-W+	1
C-LMMT35-WG	2
C-LMMT35-WM	2
```

## Aven0806001 [text]
Prompt: 请将这段brand,flavor,quantity 总结成 sku +quantity 模式
Image path: (none)
Difficulty: (not set)
Tags: (none)
Acceptance: (not set)
Notes: (none)

### Input
```text
Pulse liq 30mg
Blue Razz Ice 1
Fcuking Fab 2
Miami Mint 2
Pink Lemonade 1
Sour Strawberry 1
Triple Berries 1
Watermelon Ice 1
 
pulse liq 50mg
 
Blue Razz Ice 1
Fcuking Fab 2
Miami Mint 2
Strawberry Banana 1
```

### Expected SKU + Quantity
```text
C-PLTL30-BRI	1
C-PLTL30-FF	2
C-PLTL30-MM	2
C-PLTL30-PL	1
C-PLTL30-SS	1
C-PLTL30-TB	1
C-PLTL30-WI	1
C-PLTL50-BRI	1
C-PLTL50-FF	2
C-PLTL50-MM	2
C-PLTL50-SB	1
```

## Aven0806002 [text]
Prompt: 请将这段brand,flavor,quantity 总结成 sku +quantity 模式
Image path: (none)
Difficulty: 中等
Tags: SKU 匹配；多产品区块；数量
Acceptance: CLR、Mate Kit、Mate Pod 三个产品区块分别匹配；每行数量按输入保留。
Notes: (none)

### Input
```text
CLR 
Banana Ice 1
Blue Rancher 1 
Coffee 1 
Sour Apple Ice 1 
Sour Gush 1 
Strazz 1
White Gummy 1
 
Mate kit
Juicy Peach 1 
Miami Mint 1 
Strawberry Banana 1 
Strawberry Kiwi 1 
 
Mate pod
Blueberry Yummy 1
Miami Mint 2 
Peach Mango Watermelon 1 
Strawberry Mint Candy 1
```

### Expected SKU + Quantity
```text
C-GBCLR50-BI	1
C-GBCLR50-BR	1
C-GBCLR50-COFE	1
C-GBCLR50-SA	1
C-GBCLR50-SG	1
C-GBCLR50-STRA	1
C-GBCLR50-WG	1
C-GBMK60-JP	1
C-GBMK60-MM	1
C-GBMK60-SB	1
C-GBMK60-SK	1
C-GBMP60-BY	1
C-GBMP60-MM	2
C-GBMP60-PMW	1
C-GBMP60-SMC	1
```

## Aven0806003 [text]
Prompt: 请将这段brand,flavor,quantity 总结成 sku +quantity 模式
Image path: (none)
Difficulty: 中等
Tags: SKU 匹配；Kit/Pod；数量
Acceptance: CLIO KIT 与 CLIO POD 必须保持不同产品家族。
Notes: (none)

### Input
```text
CLIO KIT_x000B_
Fcuking Fab 1 
White Peach Raspberry 0.5
CLIO POD_x000B_
Dragonfruit Lemonade 1 
Peach Slush 1
Sour Watermelon Drop 1
```

### Expected SKU + Quantity
```text
C-GBCLK50-FF	1
C-GBCLK50-WPR	0.5
C-GBCLP50-DL	1
C-GBCLP50-PS	1
C-GBCLP50-SWD	1
```

## Aven0806004 [text]
Prompt: 请将这段brand,flavor,quantity 总结成 sku +quantity 模式
Image path: (none)
Difficulty: 困难
Tags: SKU 匹配；产品标题；New 噪声；模糊口味；人工审核
Acceptance: 忽略 (New)；Pulse 2 与 Pulse X2 必须使用不同产品家族；Blueberry Watermelon、Pink Berry Lemonade、Punch、Watermelon Crush 保留人工确认。
Notes: (none)

### Input
```text
Multiline CO: Pulse2  25k
Berry Bliss (New) 1
Mexico Mango (New) 1
Miami Mint 3
Peach Lemonade 1
Sour Strawberry 1
White Gummy Ice 1
 
 Multiline CO: PulseX 2 50k
Blackberry Blueberry  (New) 1
Blueberry Watermelon 3 
Dualicious  (New) 1
Grape Slush  (New) 6
Pink Berry Lemonade 6
Peach Slush  (New) 4
Punch 4
Raspberry Jam  (New) 1
Watermelon Crush 4
Wild Cherry Slush 1
```

### Expected SKU + Quantity
```text
C-GBP225-BB	1
C-GBP225-MM	1
C-GBP225-MMT	3
C-GBP225-PL	1
C-GBP225-SS	1
C-GBP225-WGI	1
C-GBPX250-BB	1
C-GBPX250-DUAL	1
C-GBPX250-GS	6
C-GBPX250-PS	4
C-GBPX250-RJ	1
C-GBPX250-WCS	1
```

## Aven0805001 [image_text]
Prompt: 请将这段brand,flavor,quantity 总结成 sku +quantity 模式
Image path: test_data_photo   recoginization\Screenshot 2026-08-05 at 3.23.05 PM.png
Difficulty: (not set)
Tags: (none)
Acceptance: (not set)
Notes: (none)

### Input
(No text fixture; use the image path above.)

### Expected SKU + Quantity
```text
只保留SKU 和数量即可
```

## Aven0805002 [image_text]
Prompt: 请将这段brand,flavor,quantity 总结成 sku +quantity 模式
Image path: test_data_photo recoginization\image (1).png
Difficulty: (not set)
Tags: (none)
Acceptance: (not set)
Notes: (none)

### Input
(No text fixture; use the image path above.)

### Expected SKU + Quantity
```text
(No structured expected answer)
```

## Shirley0809001 [text]
Prompt: 总结成 sku +quantity 模式
Image path: (none)
Difficulty: 中等
Tags: SKU 前置数量；负号数量；数量保留
Acceptance: 数量位于 SKU 前且带负号时取绝对值；-2.000 C-GBPX250-MMT 应输出数量 2。
Notes: (none)

### Input
```text
”-1.000	C-GBP225-BB
-1.000	C-GBP225-BW
-1.000	C-GBP225-FWG
-1.000	C-GBP225-MM
-1.000	C-GBP225-PL
-1.000	C-GBP225-SF
-1.000	C-GBP225-SWH
-1.000	C-GBP225-WPH
-1.000	C-GBPX250-BB
-1.000	C-GBPX250-GS
-2.000	C-GBPX250-MMT
-1.000	C-GBPX250-PS
-1.000	C-GBPX250-SF
-1.000	C-GBPX250-SFF
-1.000	C-GBPX250-SW
-1.000	C-GBPX25-BTF
-1.000	C-GBPX25-SFF
-1.000	C-GBPX25-WI
-1.000	C-YNGBP15K-SWD
-1.000	C-GBP15K-DSS
```

### Expected SKU + Quantity
```text
C-GBP225-BB	1
C-GBP225-BW	1
C-GBP225-FWG	1
C-GBP225-MM	1
C-GBP225-PL	1
C-GBP225-SF	1
C-GBP225-SWH	1
C-GBP225-WPH	1
C-GBPX250-BB	1
C-GBPX250-GS	1
C-GBPX250-MMT	2
C-GBPX250-PS	1
C-GBPX250-SF	1
C-GBPX250-SFF	1
C-GBPX250-SW	1
C-GBPX25-BTF	1
C-GBPX25-SFF	1
C-GBPX25-WI	1
C-YNGBP15K-SWD	1
C-GBP15K-DSS	1
```

## Shirley0810001 [text]
Prompt: 总结成 sku +quantity 模式
Image path: (none)
Difficulty: (not set)
Tags: (none)
Acceptance: (not set)
Notes: (none)

### Input
```text
OFF STAMP X-CUBE 35000 CRYSTAL KIT – 17ML
150 bars/carton

Blueberry Apple Cobbler – 1
Blackberry Blueberry – 1
Blue Razz Ice – 1
Miami Mint – 1
Mountain Berry – 1
Raspberry Watermelon – 1
Tropical Punch – 1

OFF STAMP X-CUBE 35000 CRYSTAL POD
150 bars/carton

Bangin’ Strawberry – 1
Bangin’ Watermelon – 1
Blackberry Blueberry – 1
Blueberry Sour Raspberry – 1
Cool Mint – 1
Grape – 1
Gummy Bear – 1
Maui Punch – 1
Mexican Mango – 1
Miami Mint – 2
Peach Berry – 1
Pineapple Coconut – 1
Raspberry Watermelon – 2
Rocket Freeze – 1
Scary Berry – 1
Sour Apple Ice – 1
Strawberry Cupcake – 1
Tropical Punch – 1
Watermelon Bubble Gum – 1
Watermelon Ice – 1
White Gummy – 1
Winter Mint – 1

LOST MARY NERA FULLVIEW POD
100 bars/carton

Blue Razz Ice – 1
Cool Mint – 1
Dragon Strawnana – 1
Golden Berry – 1
Miami Mint – 1
Pink & Blue – 1
Pink Lemonade – 1
Pomegranate Blast – 1
Rocket Freeze – 1
Scary Berry – 1
Watermelon Fcuking Fab – 1
```

### Expected SKU + Quantity
```text
C-XC25CSKIT-BAC	1
C-XC25CSKIT-BB	1
C-XC25CSKIT-BRI	1
C-XC25CSKIT-MM	1
C-XC25CSKIT-MB	1
C-XC25CSKIT-RW	1
C-XC25CSKIT-TP	1
C-XC25CSPOD-BS	1
C-XC25CSPOD-BW	1
C-XC25CSPOD-BB	1
C-XC25CSPOD-BSR	1
C-XC25CSPOD-CM	1
C-XC25CSPOD-GRAP	1
C-XC25CSPOD-GB	1
C-XC25CSPOD-MP	1
C-XC25CSPOD-MMG	1
C-XC25CSPOD-MM	2
C-XC25CSPOD-PB	1
C-XC25CSPOD-PC	1
C-XC25CSPOD-RW	2
C-XC25CSPOD-SB	1
C-XC25CSPOD-SAI	1
C-XC25CSPOD-SC	1
C-XC25CSPOD-TP	1
C-XC25CSPOD-WBG	1
C-XC25CSPOD-WI	1
C-XC25CSPOD-WG	1
C-XC25CSPOD-WM	1
C-LMNFP70-BRI	1
C-LMNFP70-CM	1
C-LMNFP70-DS	1
C-LMNFP70-GB	1
C-LMNFP70-MM	1
C-LMNFP70-PBLU	1
C-LMNFP70-PL	1
C-LMNFP70-PB	1
C-LMNFP70-RF	1
C-LMNFP70-SB	1
C-LMNFP70-WFF	1
C-XC25CSPOD-RF	1
```

## Pulse_20260819 [text]
Prompt: 请将这段brand,flavor,quantity 总结成 sku +quantity 模式
Image path: (none)
Difficulty: 中等
Tags: SKU匹配；产品标题简写；Pulse；数量；回归
Acceptance: Pulse 标题按 GEEK BAR PULSE 15000 目录；每行数量必须保留；18 行全部精确匹配；OMG B-Burst 不得被解析成 B-BURST 之类的 SKU。
Notes: 原始失败样本。根因：PRODUCT_HEADER_ALIASES 没有 bare Pulse，导致后续 flavor 没有产品 block；通用 SKU 正则还把 OMG B-Burst 误识别成 B-BURST。已加入 PULSE 别名并限制 prefixed SKU 必须含数字。

### Input
```text
Pulse
Berry Bliss 3
Cool Mint 8
Creamy Mintz 2
Frozen Pina Colada 5
Frozen Strawberry 2
Juicy Peach Ice 3
OMG B-Burst 5
Pink Lemonade 5
Pineapple Savers 3
Sour Apple Ice 5
Sour Blue Dust 5
Wild Berry Savers 4
White Gummy Ice 5
Blue Razz Ice 5
Blueberry Watermelon 5
Dragon Melon 5
Strawberry Banana 5
Watermelon Ice 5
```

### Expected SKU + Quantity
```text
C-GBP15K-BB	3
C-GBP15K-CMT	8
C-GBP15K-CMZ	2
C-GBP15K-FPC	5
C-GBP15K-FS	2
C-GBP15K-JPI	3
C-GBP15K-OBP	5
C-GBP15K-PL	5
C-GBP15K-PS	3
C-GBP15K-SAI	5
C-GBP15K-SBD	5
C-GBP15K-WBS	4
C-GBP15K-WGI	5
C-GBP15K-BRI	5
C-GBP15K-BW	5
C-GBP15K-DM	5
C-GBP15K-SB	5
C-GBP15K-WI	5
```

## PulseX_20260819 [text]
Prompt: 请将这段brand,flavor,quantity 总结成 sku +quantity 模式
Image path: (none)
Difficulty: 中等
Tags: SKU匹配；产品标题简写；PulseX；25K家族；数量；回归
Acceptance: PulseX 标题按 GEEK BAR PULSE X_25000 目录，不使用 Pulse X2 50000 家族；每行数量必须保留；13 行全部精确匹配；Pink&Blue 保留为一个 flavor。
Notes: 原始失败样本。根因：PRODUCT_HEADER_ALIASES 只有带完整型号的 PulseX2/Multiline 标题，没有 bare PulseX；目录对这组口味显示为标准 GEEK BAR PULSE X_25000 家族。已加入 PULSEX 别名。

### Input
```text
PulseX
Blue Rancher 5
Banana Taffy Freeze 1
Cool Mint 8
Miami mint 8
Orange Dragon 5
Orange Jam 2
Pink&Blue 1
Strawberry Colada 5
Sour Fcuking Fab 1
Strawberry Jam 1
Strawberry Kiwi Ice 2
Sour Mango Pineapple 5
Watermelon Ice 2
```

### Expected SKU + Quantity
```text
C-GBPX25-BR	5
C-GBPX25-BTF	1
C-GBPX25-CM	8
C-GBPX25-MM	8
C-GBPX25-OD	5
C-GBPX25-OJ	2
C-GBPX25-PB	1
C-GBPX25-SC	5
C-GBPX25-SFF	1
C-GBPX25-SJ	1
C-GBPX25-SKI	2
C-GBPX25-SMP	5
C-GBPX25-WI	2
```

## ShortHeaders_20260826 [text]
Prompt: 请将这段 brand、flavor、quantity 总结成 SKU + quantity 模式，并标记无法安全自动加入的行
Image path: (none)
Difficulty: 困难
Tags: SKU匹配；产品标题简写；BURJ；CLR；MATE；PulseX 2；HTML空格实体；数量；回归
Acceptance: BURJ、CLR、MATE KIT、MATE POD、PULSEX 2、PULSEX、PULSE 七个标题必须分别命中正确目录家族；`&#x20;` 只作为空白处理；38 条有明确数量的记录应精确匹配，12 条缺少数量的记录必须保留 NEEDS REVIEW，不得静默当成已确认需求。
Notes: 新失败样本。原始文本来自带 HTML 空格实体的复制内容。修复前 BURJ 未识别，`PULSEX 2` 被错误归入普通 PulseX，`&#x20;` 被当成口味行；现已增加 BURJ/PULSEX2 别名和 HTML 空白规范化。混合数量区块中没有数量的口味仍按现有安全策略进入 review（数量按 0 处理），需要用户补数量后再加入需求。

### Input
```text
BURJ
Grape Mint
Mango
&#x20;
CLR&#x20;
Watermelon Twist
&#x20;
MATE KIT
Strawberry Mint Candy
&#x20;
&#x20;
MATE POD
Amazon Lemonade
Blue Razz Ice
Cool Mint
Juicy Peach
Sour Apple Ice
Strawberry Banana
White Gummy
Watermelon Ice
&#x20;
PULSEX 2
Blackberry Blueberry
Blueberry Jam
Blue Rancher 2
Dualicious 2
Grape Slush&#x20;
Miami Mint 3
Pink&Blue
Peach Bxll
Raspberry Jam
Stone Freeze
Sour Fcuking Fab
Sour Straws&#x20;
Strawberry Watermelon 2
Wild Cherry slush&#x20;
&#x20;
&#x20;
PULSEX&#x20;
ATL Mint 4
Blackberry Blueberry 2
Cola Slush 3
Grape Slush 3
Lemon Heads 3
Miami Mint 3
Sour Apple Ice&#x20;
Strawberry Colada 3
Wild Cherry Slush 3
Watermelon Ice&#x20;
White Peach Raspberry 3
&#x20;
&#x20;
PULSE
Creamy Mintz
Dragon Melon
Frozen Strawberry
Juicy Peach Ice
OMG B-Burst
Punch
White Gummy Ice
Black Cherry
Banana Ice
Blue Razz Ice
Miami Mint
Sour Apple Ice
Sour Strawberry
```

### Expected SKU + Quantity
```text
C-GBBR80-GM	1
C-GBBR80-MANG	1
C-GBCLR50-WT	1
C-GBMK60-SMC	1
C-GBMP60-AL	1
C-GBMP60-BRI	1
C-GBMP60-CM	1
C-GBMP60-JP	1
C-GBMP60-SAI	1
C-GBMP60-SB	1
C-GBMP60-WG	1
C-GBMP60-WI	1
C-GBPX250-BR	2
C-GBPX250-DUAL	2
C-GBPX250-MMT	3
C-GBPX250-SW	2
C-GBPX25-AM	4
C-GBPX25-BB	2
C-GBPX25-CSS	3
C-GBPX25-GSS	3
C-GBPX25-LH	3
C-GBPX25-MM	3
C-GBPX25-SC	3
C-GBPX25-WCSS	3
C-GBPX25-WPR	3
C-GBP15K-CMZ	1
C-GBP15K-DM	1
C-GBP15K-FS	1
C-GBP15K-JPI	1
C-GBP15K-OBP	1
C-GBP15K-PUNC	1
C-GBP15K-WGI	1
C-GBP15K-BC	1
C-GBP15K-BI	1
C-GBP15K-BRI	1
C-GBP15K-MMI	1
C-GBP15K-SAI	1
C-GBP15K-SS	1
C-GBPX250-BB NEEDS REVIEW (缺少数量)
C-GBPX250-BJ NEEDS REVIEW (缺少数量)
C-GBPX250-GS NEEDS REVIEW (缺少数量)
C-GBPX250-PBLU NEEDS REVIEW (缺少数量)
C-GBPX250-PBX NEEDS REVIEW (缺少数量)
C-GBPX250-RJ NEEDS REVIEW (缺少数量)
C-GBPX250-SF NEEDS REVIEW (缺少数量)
C-GBPX250-SS NEEDS REVIEW (缺少数量)
C-GBPX250-SFF NEEDS REVIEW (缺少数量)
C-GBPX250-WCS NEEDS REVIEW (缺少数量)
C-GBPX25-SAI NEEDS REVIEW (缺少数量)
C-GBPX25-WI NEEDS REVIEW (缺少数量)
```

## MixedCatalog_20260908 [text]
Prompt: 请将这段订单内容总结成 SKU + quantity，并标记产品家族不明确或需要人工确认的行
Image path: (none)
Difficulty: 困难
Tags: SKU匹配；反斜杠转义；MT35000；NERA组合口味；拼写纠错；包装说明；回归
Acceptance: 初始 21 条 C-/YN SKU 行数量均保留；MT35000、NERA Kit、NERA Pod 的 33 条明确口味应匹配目录；`(100 bars/carton)` 只作为包装说明忽略；`Watermelon Fucking Fab` 应受控匹配目录的 `Watermelon Fcuking Fab`；`Mint Lemonade` 因标题未声明 Lemonade Edition，必须保留 NEEDS REVIEW。
Notes: 新失败样本。修复前 `Watermelon Fucking Fab` 被模糊匹配并省略，两个包装说明行被当成口味 review；现已增加受控拼写别名并忽略独立包装说明。`Mint Lemonade` 在目录中属于 `LOST MARY (LEMONADE EDITION)_MT35000`，而原始标题只有普通 `LOST MARY MT35000`，不能在没有产品版本信息时自动跨家族指定 SKU。

### Input
```text
C-GBBR80-BG	1.000\
C-GBBR80-BR	1.000\
C-GBBR80-LM	1.000\
C-GBBR80-MANG	1.000\
C-GBBR80-PI	1.000\
C-GBMK60-AL	1.000\
C-GBMK60-BS	1.000\
C-GBMK60-BY	1.000\
C-GBMK60-JP	1.000\
C-GBMK60-PMW	1.000\
C-GBMK60-SMC	1.000\
C-GBMK60-WG	1.000\
C-GBMP60-BY	1.000\
C-GBMP60-JP	1.000\
C-GBMP60-SMC	1.000\
C-GBPX25-BRI	1.000\
C-GBPX25-MM	1.000\
C-GBPX25-SFF	1.000\
C-GBPX25-WCSS	1.000\
C-YNGBP15K-SB	1.000\
C-GBP15K-WBS	1.000\
LOST MARY MT35000 18ML

Black Mint+ – 1
Blackberry Blueberry+ – 1
Classic Tobacco – 2
Clear – 1
Kiwi Passion Fruit – 1
Miami Mint – 1
Mountain Berry – 1
Strawberry Kiwi+ – 1
Strawberry+ – 1
Toasted Banana – 1
White Gami+ – 1
Winter Mint – 2
Purple Pixy – 1
White Pixy – 1
Mint Lemonade – 1
LOST MARY NERA FULLVIEW KIT 70000 24ML
(100 bars/carton)

Blue Razz Ice – 1
Baja Splash & Hawaiian Punch – 1
Clear – 1
Juicy Peach Ice & Pineapple Coconut – 1
Miami Mint & Winter Mint – 1
Pink Lemonade & Pink & Blue – 1
Scary Berry & Golden Berry – 1
Sour Apple Ice – 1
Strawberry Ice & Strawberry Watermelon – 1

LOST MARY NERA FULLVIEW POD
(100 bars/carton)

Baja Splash – 1
Blue Razz Ice – 1
Cool Mint – 1
Hawaiian Punch – 1
Juicy Peach Ice – 1
Miami Mint – 1
Pineapple Coconut – 1
Strawberry Watermelon – 1
Watermelon Fucking Fab – 1
```

### Expected SKU + Quantity
```text
C-GBBR80-BG	1
C-GBBR80-BR	1
C-GBBR80-LM	1
C-GBBR80-MANG	1
C-GBBR80-PI	1
C-GBMK60-AL	1
C-GBMK60-BS	1
C-GBMK60-BY	1
C-GBMK60-JP	1
C-GBMK60-PMW	1
C-GBMK60-SMC	1
C-GBMK60-WG	1
C-GBMP60-BY	1
C-GBMP60-JP	1
C-GBMP60-SMC	1
C-GBPX25-BRI	1
C-GBPX25-MM	1
C-GBPX25-SFF	1
C-GBPX25-WCSS	1
C-YNGBP15K-SB	1
C-GBP15K-WBS	1
C-LMMT35-BM+	1
C-LMMT35-BB+	1
C-LMMT35-CT	2
C-LMMT35-CLEA	1
C-LMMT35-KPF	1
C-LMMT35-MM	1
C-LMMT35-MB	1
C-LMMT35-SK+	1
C-LMMT35-SBER+	1
C-LMMT35-TB	1
C-LMMT35-WG	1
C-LMMT35-WM	2
C-LMMT35-PP	1
C-LMMT35-WP	1
C-LMNFK70-BRI	1
C-LMNFK70-BSHP	1
C-LMNFK70-CLEA	1
C-LMNFK70-JPPC	1
C-LMNFK70-MMWM	1
C-LMNFK70-PLPB	1
C-LMNFK70-SBGB	1
C-LMNFK70-SAI	1
C-LMNFK70-SISW	1
C-LMNFP70-BS	1
C-LMNFP70-BRI	1
C-LMNFP70-CM	1
C-LMNFP70-HP	1
C-LMNFP70-JPI	1
C-LMNFP70-MM	1
C-LMNFP70-PC	1
C-LMNFP70-SW	1
C-LMNFP70-WFF	1
C-LMMT35LE-ML NEEDS REVIEW (标题未声明 Lemonade Edition)
```

## OffStampCrystalPod_20260908 [text]
Prompt: 请将这段产品标题和口味数量转换成 SKU + quantity
Image path: (none)
Difficulty: 中等
Tags: SKU匹配；Off Stamp；Crystal Pod；反斜杠转义；连字符数量；拼写纠错；回归
Acceptance: `Off stamp crystal pod` 必须匹配 `OFF STAMP X-CUBE POD (CRYSTAL)_35000`；行尾反斜杠必须忽略；6 条口味全部精确匹配，数量合计 7。
Notes: 原始文本每行末尾带 `\`，并使用 `flavor-quantity` 格式。修复前反斜杠阻止数量正则识别，整组无法生成有效需求；现已将行尾反斜杠作为复制格式噪声清除。`Bangun strawberry` 和 `Banging watermelon` 是目录 `Bangin' Strawberry`、`Bangin' Watermelon` 的受控拼写变体。

### Input
```text
Off stamp crystal pod\
Winter mint-1\
Bangun strawberry-1\
Blue razz grape ice-1\
Tropical punch-1\
Miami mint-2\
Banging watermelon-1
```

### Expected SKU + Quantity
```text
C-XC25CSPOD-WM	1
C-XC25CSPOD-BS	1
C-XC25CSPOD-BRGI	1
C-XC25CSPOD-TP	1
C-XC25CSPOD-MM	2
C-XC25CSPOD-BW	1
```
