; Copyright 2021 Satoru NAKAYA
;
; Licensed under the Apache License, Version 2.0 (the "License");
; you may not use this file except in compliance with the License.
; You may obtain a copy of the License at
;
;	  http://www.apache.org/licenses/LICENSE-2.0
;
; Unless required by applicable law or agreed to in writing, software
; distributed under the License is distributed on an "AS IS" BASIS,
; WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
; See the License for the specific language governing permissions and
; limitations under the License.


; **********************************************************************
; 単純トリル法による連接速度を計測
;	出力は 1.05 秒後にまとめて
; http://oookaworks.seesaa.net/article/490449134.html#gsc.tab=0
; **********************************************************************

; ----------------------------------------------------------------------
; 初期設定
; ----------------------------------------------------------------------

SetWorkingDir %A_ScriptDir%	; スクリプトの作業ディレクトリを変更
#SingleInstance force		; 既存のプロセスを終了して実行開始
#Persistent					; スクリプトを常駐状態にする
#NoEnv						; 変数名を解釈するとき、環境変数を無視する
SetBatchLines, -1			; 自動Sleepなし
ListLines, Off				; スクリプトの実行履歴を取らない
SetKeyDelay, -1, -1			; キーストローク間のディレイを変更
#MenuMaskKey vk07			; Win または Alt の押下解除時のイベントを隠蔽するためのキーを変更する
#UseHook					; ホットキーはすべてフックを使用する
;Process, Priority, , High	; プロセスの優先度を変更
Thread, interrupt, 15, 17	; スレッド開始から15ミリ秒ないし17行以内の割り込みを、絶対禁止
;SetStoreCapslockMode, off	; Sendコマンド実行時にCapsLockの状態を自動的に変更しない

;SetFormat, Integer, H		; 数値演算の結果を、16進数の整数による文字列で表現する

#HotkeyInterval 200			; 指定時間(ミリ秒単位)の間に実行できる最大のホットキー数
#MaxHotkeysPerInterval 200	; 指定時間の間に実行できる最大のホットキー数

; ----------------------------------------------------------------------
; グローバル変数
; ----------------------------------------------------------------------

; 設定
passCount := 40		; Int型定数		この個数のキーを押したところまでの時間も出力させる

; 入力バッファ
changedKeys := []	; [String]型
changedTimes := []	; [Double]型	入力の時間
pressedKeys := []	; [String]型

repeatKeyCount := 0	; Int型
trillCount := 0		; Int型
trillError := False	; Bool型
firstPressTime :=	; Double?型
nowKeyTime := 0.0	; Double型
;nowKeyName			; String型
lastKeyName :=		; String?型
;clipSaved :=

; キーボードドライバを調べて keyDriver に格納する
; 参考: https://ixsvr.dyndns.org/blog/764
RegRead, keyDriver, HKEY_LOCAL_MACHINE, SYSTEM\CurrentControlSet\Services\i8042prt\Parameters, LayerDriver JPN
		; keyDriver: String型

If (keyDriver = "kbd101.dll")
	scArray := ["Esc", "1", "2", "3", "4", "5", "6", "7", "8", "9", "Ø", "-", "=", "BackSpace", "Tab"
		, "Q", "W", "E", "R", "T", "Y", "U", "I", "O", "P", "[", "]", "", "", "A", "S"
		, "D", "F", "G", "H", "J", "K", "L", ";", "'", "`", "LShift", "＼", "Z", "X", "C", "V"
		, "B", "N", "M", ",", ".", "/", "", "", "", "Space", "CapsLock", "F1", "F2", "F3", "F4", "F5"
		, "F6", "F7", "F8", "F9", "F10", "Pause", "ScrollLock", "", "", "", "", "", "", "", "", ""
		, "", "", "", "", "SysRq", "", "KC_NUBS", "F11", "F12", "(Mac)=", "", "", "(NEC),", "", "", ""
		, "", "", "", "", "F13", "F14", "F15", "F16", "F17", "F18", "F19", "F20", "F21", "F22", "F23", ""
		, "(JIS)ひらがな", "(Mac)英数", "(Mac)かな", "(JIS)_", "", "", "F24", "KC_LANG4"
		, "KC_LANG3", "(JIS)変換", "", "(JIS)無変換", "", "(JIS)￥", "(Mac),", ""]
				; [String]型
Else
	scArray := ["Esc", "1", "2", "3", "4", "5", "6", "7", "8", "9", "Ø", "-", "^", "BackSpace", "Tab"
		, "Q", "W", "E", "R", "T", "Y", "U", "I", "O", "P", "@", "[", "", "", "A", "S"
		, "D", "F", "G", "H", "J", "K", "L", ";", ":", "半角/全角", "LShift", "]", "Z", "X", "C", "V"
		, "B", "N", "M", ",", ".", "/", "", "", "", "Space", "英数", "F1", "F2", "F3", "F4", "F5"
		, "F6", "F7", "F8", "F9", "F10", "Pause", "ScrollLock", "", "", "", "", "", "", "", "", ""
		, "", "", "", "", "SysRq", "", "KC_NUBS", "F11", "F12", "(Mac)=", "", "", "(NEC),", "", "", ""
		, "", "", "", "", "F13", "F14", "F15", "F16", "F17", "F18", "F19", "F20", "F21", "F22", "F23", ""
		, "(JIS)ひらがな", "(Mac)英数", "(Mac)かな", "(JIS)_", "", "", "F24", "KC_LANG4"
		, "KC_LANG3", "(JIS)変換", "", "(JIS)無変換", "", "(JIS)￥", "(Mac),", ""]
				; [String]型

; ----------------------------------------------------------------------
; 起動
; ----------------------------------------------------------------------

	Run, Notepad.exe, , , pid	; メモ帳を起動
	Sleep, 500
	WinActivate, ahk_pid %pid%	; アクティブ化
	If (A_IsCompiled)
	{
		; 実行ファイル化されたスクリプトの場合は終了
		Send, 実行ファイル化されているので終了します。
		ExitApp
	}
	Clipboard := "単純トリル法による連接速度を計測します。Escキーを押すと終了します。`n`n"
		. SeparateLines() . "`n"
	Send, ^v

Exit	; 起動時はここまで実行

; ----------------------------------------------------------------------
; タイマー関数、設定
; ----------------------------------------------------------------------

; 参照: https://www.autohotkey.com/boards/viewtopic.php?t=36016
QPCInit() {	; () -> Int64
	DllCall("QueryPerformanceFrequency", "Int64P", freq)	; freq: Int64型
	Return freq
}
QPC() {		; () -> Double	ミリ秒単位
	static coefficient := 1000.0 / QPCInit()	; Double型
	DllCall("QueryPerformanceCounter", "Int64P", count)	; count: Int64型
	Return count * coefficient
}

; ----------------------------------------------------------------------
; サブルーチン
; ----------------------------------------------------------------------

OutputTimer:
	Output()
	; 変数のリセット
	firstPressTime :=
	pressedKeys := []
	repeatKeyCount := 0
	trillCount := 0
	trillError := False
	Return

; ----------------------------------------------------------------------
; 関数
; ----------------------------------------------------------------------

SeparateLines()	; () -> String
{
	global passCount
;	local str	; String型
;		, i		; Int型
	str := "", i = 1
	While (i <= passCount)
	{
		If (Mod(i, 10) == 0)
			str .= "*"
		Else If (Mod(i, 5) == 0)
			str .= "+"
		Else
			str .= "-"
		i++
	}
	Return str
}

Output()	; () -> Double?
{
	global pid, changedKeys, changedTimes, scArray, passCount, trillError
	static lastKeyTime := QPC()		; Double型
;	local keyName, postKeyName, lastPostKeyName, temp		; String型
;		, outputString										; String型
;		, keyTime, startTime								; Double型
;		, firstPressTime, passTime							; Dount?型
;		, pressKeyCount, releaseKeyCount, repeatKeyCount, 	; Int型
;		, i, number, multiPress		; Int型
;		, pressingKeys				; [String]型

	; 変数の初期化
	pressKeyCount := repeatKeyCount := releaseKeyCount := 0
	multiPress := 0
	firstPressTime :=
	passTime :=
	pressingKeys := []
	outputString := "`n" . SeparateLines()
	; 起動から、または前回表示からの経過時間表示が必要なら次の初期値は " " とする
	lastPostKeyName := ""

	; 一塊の入力の先頭の時間を保存
	startTime := changedTimes[1]

	; 入力バッファが空になるまで
	While (changedKeys.Length())
	{
		; 入力バッファから読み出し
		keyName := changedKeys.RemoveAt(1), keyTime := changedTimes.RemoveAt(1)
		StringTrimLeft, keyName, keyName, 1		; 頭の ~ を取り除く

		; キーの上げ下げを調べる
		StringRight, postKeyName, keyName, 3	; postKeyName に入力末尾の2文字を入れる
		; キーが離されたとき
		If (postKeyName = " up")
		{
			StringTrimRight, keyName, keyName, 3
			releaseKeyCount++
			; ロールオーバー押し検出用 押しているキーを入れた配列から消す
			i := 1
			While (i <= pressingKeys.Length())
			{
				If (keyName = pressingKeys[i])
				{
					pressingKeys.RemoveAt(i)
					Break
				}
				i++
			}

			preKeyName := "", postKeyName := "↑"
			If (lastPostKeyName != postKeyName)
				outputString .= "`n`t`t"	; キーの上げ下げが変わったら改行と字下げ
			Else
				outputString .= " "
		}
		Else
		{
			If (!firstPressTime)
				firstPressTime := keyTime
			; キーリピートでないキーを数える
			If (keyName != pressingKeys[pressingKeys.Length()])
			{
				pressKeyCount++
				preKeyName := "", postKeyName := "↓"
			}
			Else
			{
				repeatKeyCount++
				preKeyName := "<", postKeyName := ">"
			}
			; ロールオーバー押し検出 押しているキーを入れた配列と比べる
			i := 1
			While (i <= pressingKeys.Length())
			{
				If (keyName = pressingKeys[i])
					Break
				i++
			}
			If (i > pressingKeys.Length())
			{
				; 配列に追加
				pressingKeys.Push(keyName)
				; 同時押し数更新
				If (i > multiPress)
					multiPress := i
			}

			; 設定した個数なら時間を保存
			If (!trillError && !repeatKeyCount && pressKeyCount == passCount)
				passTime := keyTime

			If (lastPostKeyName != "↓" && lastPostKeyName != ">")
				outputString .= "`n"	; キーの上げ下げが変わったら改行
			Else
				outputString .= " "
		}
		; 前回の入力からの時間を書き出し
		If (lastPostKeyName != "")
			outputString .= "(" . Round(keyTime - lastKeyTime, 1) . "ms) "

		; 入力文字の書き出し
		If (keyName = "LWin")		; LWin を半角のまま出力すると、なぜか Win+L が発動する
			keyName := "ＬWin"
		Else If (keyName = "vk1A")
			keyName := "(Mac)英数"
		Else If (keyName = "vk16")
			keyName := "(Mac)かな"
		Else
		{
			If (SubStr(keyName, 1, 2) = "sc")
			{
				number := "0x" . SubStr(keyName, 3, 2)
				temp := scArray[number]
				If (temp != "")
					keyName := temp
			}
		}
		outputString .= preKeyName . keyName . postKeyName

		; 変数の更新
		lastKeyTime := keyTime	; 押した時間を保存
		lastPostKeyName := postKeyName		; キーの上げ下げを保存
	}

	; 一塊の入力時間合計を出力
	outputString .= "`n***** キー変化 " . pressKeyCount + repeatKeyCount + releaseKeyCount
		. " 回で " . Round(keyTime - startTime, 1) . "ms。`n`t("
		. pressKeyCount . " 個押し + " . repeatKeyCount . " 個キーリピート + " . releaseKeyCount . " 個離す)`n"
	If (multiPress > 1)
		outputString .= "`t同時押し 最高 " . multiPress . " キー。`n"
	If (passTime)
		outputString .= "`t" . passCount . " 個目を押すまでに "
			. Round((passTime - firstPressTime) / 1000, 3) . " 秒。`n"
	If (trillError)
		outputString .= "`t繰り返しが乱れました。`n"
	outputString .= "`n`n" . SeparateLines() . "`n"
	Clipboard := outputString

	; 先ほど起動したメモ帳にのみ出力
	IfWinActive, ahk_pid %pid%
		IfWinNotActive , ahk_class #32770
			Send, ^v

	Return
}

; ----------------------------------------------------------------------
; ホットキー
;		コメントの中で、:: がついていたら down と up のセットで入れ替え可能
; ※キーの調査には、ソフトウェア Keymill Ver.1.4 を使用しました。
;		http://kts.sakaiweb.com/keymill.html
; ※参考：https://docs.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes
; ----------------------------------------------------------------------
Esc::
	ExitApp

#MaxThreadsPerHotkey 3	; 1つのホットキー・ホットストリングに多重起動可能な
						; 最大のスレッド数を設定

; キー入力部
~sc29::	; (JIS)半角/全角	(US)`
~sc02::		; 1::	vk31::
~sc03::		; 2::	vk32::
~sc04::		; 3::	vk33::
~sc05::		; 4::	vk34::
~sc06::		; 5::	vk35::
~sc07::		; 6::	vk36::
~sc08::		; 7::	vk37::
~sc09::		; 8::	vk38::
~sc0A::		; 9::	vk39::
~sc0B::		; 0::	vk30::
~sc0C::		; -::	vkBD::
~sc0D::	; (JIS)^	(US)=
~sc7D::	; (JIS)￥
~sc10::		; Q::	vk51::
~sc11::		; W::	vk57::
~sc12::		; E::	vk45::
~sc13::		; R::	vk52::
~sc14::		; T::	vk54::
~sc15::		; Y::	vk59::
~sc16::		; U::	vk55::
~sc17::		; I::	vk49::
~sc18::		; O::	vk4F::
~sc19::		; P::	vk50::
~sc1A::	; (JIS)@	(US)[
~sc1B::	; (JIS)[	(US)]
~sc56::	; KC_NUBS
~sc1E::		; A::	vk41::
~sc1F::		; S::	vk53::
~sc20::		; D::	vk44::
~sc21::		; F::	vk46::
~sc22::		; G::	vk47::
~sc23::		; H::	vk48::
~sc24::		; J::	vk4A::
~sc25::		; K::	vk4B::
~sc26::		; L::	vk4C::
~sc27::		; `;::
~sc28::	; (JIS):	(US)'
~sc2B::	; (JIS)]	(US)＼
~sc2C::		; Z::	vk5A::
~sc2D::		; X::	vk58::
~sc2E::		; C::	vk43::
~sc2F::		; V::	vk56::
~sc30::		; B::	vk42::
~sc31::		; N::	vk4E::
~sc32::		; M::	vk4D::
~sc33::	; ,			vkBC::
~sc34::		; .::	vkBE::
~sc35::		; /::	vkBF::
~sc73::	; (JIS)_
~sc39::		; Space::	vk20::
~sc79::		; (JIS)変換
~sc7B::		; (JIS)無変換
	; 入力バッファへ保存
	changedKeys.Push(nowKeyName := A_ThisHotkey), changedTimes.Push(nowKeyTime := QPC())
	pressedKeys.Push(nowKeyName)
	If (!firstPressTime)
		firstPressTime := nowKeyTime
	; キー変化なく1.05秒たったら表示
	SetTimer, OutputTimer, -1050
	; 先ほど起動したメモ帳のときだけ後の判定をする
	IfWinNotActive, ahk_pid %pid%
		Return
	; キーリピート検出
	If (nowKeyName == lastKeyName)
		repeatKeyCount++

	; 繰り返しパターンの判定 1文字目
	If (!trillCount)
		trillCount--	; 1周目は負数でカウント
	; 繰り返しパターン 1周目2文字目以降
	Else If (trillCount < 0)
	{
		; 1文字目と同じになるまでカウントする
		If (pressedKeys[1] != nowKeyName)
			trillCount--
		; 2周目に入った
		Else
			trillCount := - trillCount	; 正数に直す
	}
	; 繰り返しパターン 2周目以降 繰り返しが乱れたか
	Else If (!trillError && pressedKeys[pressedKeys.Length() - trillCount] != nowKeyName
		&& pressedKeys.Length() <= passCount)
	{
		trillError := True
		TrayTip, , 繰り返しが乱れました
	}

	; 設定の個数に達した
	If (!trillError && !repeatKeyCount && pressedKeys.Length() == passCount)
		TrayTip, , % passCount . " 個目を押すまでに "
			. Round((nowKeyTime - firstPressTime) / 1000, 3)
			. " 秒"
	; 変数の更新
	lastKeyName := nowKeyName
	Return


; キー押上げ
~sc29 up::	; (JIS)半角/全角	(US)`
~sc02 up::		; 1 up::	vk31 up::
~sc03 up::		; 2 up::	vk32 up::
~sc04 up::		; 3 up::	vk33 up::
~sc05 up::		; 4 up::	vk34 up::
~sc06 up::		; 5 up::	vk35 up::
~sc07 up::		; 6 up::	vk36 up::
~sc08 up::		; 7 up::	vk37 up::
~sc09 up::		; 8 up::	vk38 up::
~sc0A up::		; 9 up::	vk39 up::
~sc0B up::		; 0 up::	vk30 up::
~sc0C up::		; - up::	vkBD up::
~sc0D up::	; (JIS)^	(US)=
~sc7D up::	; (JIS)￥
~sc10 up::		; Q up::	vk51 up::
~sc11 up::		; W up::	vk57 up::
~sc12 up::		; E up::	vk45 up::
~sc13 up::		; R up::	vk52 up::
~sc14 up::		; T up::	vk54 up::
~sc15 up::		; Y up::	vk59 up::
~sc16 up::		; U up::	vk55 up::
~sc17 up::		; I up::	vk49 up::
~sc18 up::		; O up::	vk4F up::
~sc19 up::		; P up::	vk50 up::
~sc1A up::	; (JIS)@	(US)[
~sc1B up::	; (JIS)[	(US)]
~sc56 up::	; KC_NUBS
~sc1E up::		; A up::	vk41 up::
~sc1F up::		; S up::	vk53 up::
~sc20 up::		; D up::	vk44 up::
~sc21 up::		; F up::	vk46 up::
~sc22 up::		; G up::	vk47 up::
~sc23 up::		; H up::	vk48 up::
~sc24 up::		; J up::	vk4A up::
~sc25 up::		; K up::	vk4B up::
~sc26 up::		; L up::	vk4C up::
~sc27 up::		; `; up::
~sc28 up::	; (JIS):	(US)'
~sc2B up::	; (JIS)]	(US)＼
~sc2C up::		; Z up::	vk5A up::
~sc2D up::		; X up::	vk58 up::
~sc2E up::		; C up::	vk43 up::
~sc2F up::		; V up::	vk56 up::
~sc30 up::		; B up::	vk42 up::
~sc31 up::		; N up::	vk4E up::
~sc32 up::		; M up::	vk4D up::
~sc33 up::	; ,				vkBC up::
~sc34 up::		; . up::	vkBE up::
~sc35 up::		; / up::	vkBF up::
~sc73 up::	; (JIS)_
~sc39 up::		; Space up::	vk20 up::
~sc79 up::		; (JIS)変換
~sc7B up::		; (JIS)無変換
	; 入力バッファへ保存
	changedKeys.Push(A_ThisHotkey), changedTimes.Push(QPC())
	SetTimer, OutputTimer, -1050	; キー変化なく1.05秒たったら表示
	If (A_ThisHotkey = nowKeyName . " up")
		lastKeyName :=
	Return

#MaxThreadsPerHotkey 1	; 元に戻す
