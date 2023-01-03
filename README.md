# auto-set-notes
Automatically set notes (myanmar kyats) in  Microsoft Excel

သတ်မှတ်ပေးသော တန်ဖိုး (ပမာဏ) ပေါ် မူတည်၍ သင့်တော်သော မြန်မာကျပ်တန်များ အရေအတွက်  မည်သို့ ပေးသင့်ပုံကို ပြန်ပြပေးသော Excel တစ်ခု ဖြစ်ပါသည်။

<img src="https://github.com/chitkokooo/excel-figure-to-word-en-mm/blob/main/hidden_var.png" alt="hidden_var excel Sheet">

## လိုအပ်ချက်များ
- Macro-Enabled Excel File များ အလုပ်လုပ်နိုင်သော အခြေအနေ ရှိရမည်။
- Macro-Enabled File `.xlsm` အနေဖြင့် ဖိုင်ကို သိမ်းဆည်းပေးရန် လိုအပ်သည်။
- နမူနာဖိုင်တွင် `MS Excel Version 2010` နှင့် `Pyidaungsu Book Font` ကို အသုံးပြုထားပါသည်။

##  Coding အတွက် ကြိုတင်ပြင်ဆင်မှု
- အသုံးပြုလိုသော Sheet Name ကို `Line 23` ရှိ `"Notes"` နေရာတွင် အစားထိုး ပြင်ဆင်ပါ။
- `Line 26` မှ `Line 37` အထိ Cell များကို Notes အမျိုးအစား ခွဲခြားသတ်မှတ်ရန် မိမိနှစ်သက်ရာ Cell ဖြင့် အစားထိုး ပြင်ဆင်ပါ။
```
    ' -------------------- Variable Declaration ------------------------
    ' In this section, you can change sheet_name (Notes) and note locations (cell numbers)
    
    ' Sheet Name
    Dim sheet_name As String
    sheet_name = "Notes"
    
    ' Cell positions for Notes
    notes_10000 = "D5"
    notes_5000 = "D6"
    notes_1000 = "D7"
    notes_500 = "D8"
    notes_200 = "D9"
    notes_100 = "D10"
    notes_50 = "D11"
    notes_20 = "D12"
    notes_10 = "D13"
    notes_5 = "D14"
    notes_1 = "D15"
    coin = "E16"
    ' ------------------- End of Variable Declaration -----------------------
```

## Resources
