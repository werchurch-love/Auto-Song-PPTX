# 🎵 Auto-Song-PPTX  
## 一鍵製作多張詩歌投影片

在 `lyrics.txt` 寫歌詞，自動產生多首 `.pptx`。  
自動分頁、背景圖及段落跳轉按鈕。✅

▶️ **[觀看示範影片](https://youtu.be/irIsSUxqdDs?si=nKD05K_P_RvnZoT5)**

***

## ✅ 使用前

- Windows 電腦
- PowerPoint 桌面版
- 先儲存及關閉 PowerPoint

> [!IMPORTANT]
> Mac、手機、網頁版 PowerPoint 不支援。

***

## 📥 第一次設定

1. GitHub：**Code → Download ZIP**
2. 解壓縮；以下檔案必須放在同一資料夾：

```text
create-lyrics-ppt.vbs
lyrics.txt
background.jpg
```

`background-1.jpg`、`background-2.jpg` 為可選背景。

***

## 📝 寫歌詞

開啟 `lyrics.txt`：

```text
[在祢沒有難成的事]
60
dark blue

=1=
芥菜種的信心 可以將大山挪開
相信神的應許 可以跨越困難

=C1=
在祢沒有 沒有難成的事
軟弱和憂慮中 祢不曾放棄我
```

| 寫法 | 意思 |
|---|---|
| `[歌名]` | 新歌 |
| `=1=`、`=C1=` | 新段落＋按鈕 |
| `==` | 只分頁 |

> [!CAUTION]
> 歌詞標點會自動變成空格。

***

## 🚀 製作方法

1. 雙擊 `create-lyrics-ppt.vbs`
2. 有提示時輸入背景編號；留空即用 `background.jpg`
3. 等待完成 ✅

每首歌會生成一個 `.pptx`。

> [!CAUTION]
> 同名 `.pptx` 會被直接覆蓋。手動修改的檔案請先改名或搬走。

***

## 🔐 安全

`.vbs` 雙擊會立即執行，只應使用可信來源的檔案。

查看程式碼時，**不要雙擊**：

1. 右鍵按 `.vbs`
2. 「開啟檔案方式」→「記事本」

不要把密碼、個人資料或內部資料貼到公開 AI。

***

授權：GPL-3.0
