Sub wpBulk()
Dim bot As New WebDriver
Dim snd As New Keys

bot.Start "chrome", "https://web.whatsapp.com/"
bot.Get "/"

MsgBox "Devam etmek için QR Kod okutmanız gerekmektedir."
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow
    arakisi = Sheets(1).Range("A" & i).Value
    msjgonder = Sheets(1).Range("B" & i).Value
    bot.FindElementByXPath("//*[@id='side']/div[1]/div/label/div/div[2]").Click
    bot.Wait (500)
    bot.SendKeys (arakisi)
    bot.Wait (500)
    bot.SendKeys (snd.Enter)
    bot.Wait (500)
    bot.SendKeys (msjgonder)
    bot.Wait (500)
    bot.SendKeys (snd.Enter)
Next i
MsgBox "Gönderilmiştir :)"
Stop


End Sub


