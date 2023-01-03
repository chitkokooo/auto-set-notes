# auto-set-notes
Automatically set notes (myanmar kyats) in  Microsoft Excel

သတ်မှတ်ပေးသော တန်ဖိုး (ပမာဏ) ပေါ် မူတည်၍ သင့်တော်သော မြန်မာကျပ်တန်များ အရေအတွက်  မည်သို့ ပေးသင့်ပုံကို ပြန်ပြပေးသော Excel တစ်ခု ဖြစ်ပါသည်။

<img src="https://github.com/chitkokooo/auto-set-notes/blob/main/demo.gif" alt="excel auto set note demo">

## လိုအပ်ချက်များ
- Macro-Enabled Excel File များ အလုပ်လုပ်နိုင်သော အခြေအနေ ရှိရမည်။
- Macro-Enabled File `.xlsm` အနေဖြင့် ဖိုင်ကို သိမ်းဆည်းပေးရန် လိုအပ်သည်။
- နမူနာဖိုင်တွင် `MS Excel Version 2010` နှင့် `Pyidaungsu Book Font` ကို အသုံးပြုထားပါသည်။

##  Coding အတွက် ကြိုတင်ပြင်ဆင်မှု
- အသုံးပြုလိုသော Sheet Name ကို `Line 23` ရှိ `"Notes"` နေရာတွင် အစားထိုး ပြင်ဆင်ပါ။
- `Line 26` မှ `Line 37` အထိ Cell များကို Notes အမျိုးအစား ခွဲခြားသတ်မှတ်ရန် မိမိနှစ်သက်ရာ Cell ဖြင့် အစားထိုး ပြင်ဆင်ပါ။
```
Line 18    ' -------------------- Variable Declaration ------------------------
Line 19    ' In this section, you can change sheet_name (Notes) and note locations (cell numbers)
Line 20    
Line 21    ' Sheet Name
Line 22    Dim sheet_name As String
Line 23    sheet_name = "Notes"
Line 24    
Line 25    ' Cell positions for Notes
Line 26    notes_10000 = "D5"
Line 27    notes_5000 = "D6"
Line 28    notes_1000 = "D7"
Line 29    notes_500 = "D8"
Line 30    notes_200 = "D9"
Line 31    notes_100 = "D10"
Line 32    notes_50 = "D11"
Line 33    notes_20 = "D12"
Line 34    notes_10 = "D13"
Line 35    notes_5 = "D14"
Line 36    notes_1 = "D15"
Line 37    coin = "E16"
LIne 38    ' ------------------- End of Variable Declaration -----------------------
```
<img src="https://github.com/chitkokooo/auto-set-notes/blob/main/settings.gif" alt="excel auto set note settings">

## Resources
