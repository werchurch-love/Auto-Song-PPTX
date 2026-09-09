# 🎵 Auto-Song-PPTX

### 一鍵製作多張詩歌投影片

在 `lyrics.txt` 寫歌詞，自動產生多首 `.pptx`。  
自動分頁、背景圖及段落跳轉按鈕。✅

▶️ **[觀看示範影片](https://youtu.be/irIsSUxqdDs?si=nKD05K_P_RvnZoT5)**

##

### 🔐 安全

> [!TIP]
> `.vbs` 是 Windows 自動化程式，雙擊會立即執行。

#### 如何檢視程式碼

1. 右鍵按 `.vbs`
2. 選「開啟檔案方式」
3. 選「記事本」

#### 使用 AI 檢查

想知程式碼是否安全，可將按以上方式直接檢視。或者將整個 vbs 文件貼到你信任的 AI，並詢問：

```text
請用繁體中文解釋這段 VBS 的功能。

請檢查它有沒有刪除、覆蓋、移動或複製檔案，
執行 CMD／PowerShell、下載資料、連接網絡、
讀取密碼或上傳資料。

請列出它會讀取、建立、修改或刪除的檔案和資料夾。
請不要修改程式碼，只報告實際功能和風險。
```

> [!CAUTION]
> 不要把密碼、個人資料或內部檔案貼到公開 AI。

##

### ✅ 使用前

- Windows 電腦
- PowerPoint 桌面版
- 先儲存及關閉 PowerPoint

> [!IMPORTANT]
> Mac、手機、網頁版 PowerPoint 不支援。

##

### 📥 第一次設定

1. GitHub：**Code → Download ZIP**
2. 解壓縮；以下檔案必須放在同一資料夾：

```text
create-lyrics-ppt.vbs
lyrics.txt
background.jpg
```

> [!TIP]
>
> - `background-1.jpg`、`background-2.jpg` 等可作批量背景
> - 與歌曲同名圖片可作指定背景，例如 `神大愛.jpg`

##

### 📝 寫歌詞

開啟 `lyrics.txt`：

```text
[在祢沒有難成的事]
60pt
dark blue
left / center / right

=1=
芥菜種的信心 可以將大山挪開
相信神的應許 可以跨越困難

=C1=
在祢沒有 沒有難成的事
軟弱和憂慮中 祢不曾放棄我
```

> [!IMPORTANT]
>
> - 歌名必須用中括號包住 `[] ` 或 `【】`
> - 歌詞的全部標點符號會自動變成空格

| 寫法          | 意思         |
| ------------- | ------------ |
| `[歌名]`      | 新歌         |
| `=1=`、`=C1=` | 新段落＋按鈕 |
| `==`          | 只分頁       |

##

### 🚀 製作方法

1. 雙擊 `create-lyrics-ppt.vbs`
2. 有提示時輸入背景編號；留空即用 `background.jpg`
3. 等待完成 ✅

每首詩歌均會生成一個獨立 `.pptx`。

> [!CAUTION]
> 資料夾內同名 `.pptx` 會被直接覆蓋。手動修改後的檔案請先改名或搬走。

---

授權：GPL-3.0
