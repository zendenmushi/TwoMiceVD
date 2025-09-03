# TwoMiceVD プロジェクト構造

## 概要
2台のマウスで仮想デスクトップを切り替えるWindows 11アプリケーションのプロジェクト構造です。1206行の単一ファイルから論理的な13個のファイルに分割し、責任分離の原則に基づいて整理されています。

## フォルダー構造

```
TwoMiceVD/
├── Program.cs                      # アプリケーションエントリポイント
├── Core/                           # コア機能・ビジネスロジック
│   ├── Events.cs                   # DeviceMovedEventArgs (共通イベント引数)
│   ├── VirtualDesktopController.cs # 仮想デスクトップ制御 (InputSimulator使用)
│   └── SwitchPolicy.cs             # 切り替えポリシー (閾値・ヒステリシス管理)
├── Input/                          # 入力処理レイヤー
│   └── RawInputManager.cs          # Raw Input API処理 (マウス識別・イベント発生)
├── UI/                             # ユーザーインターフェース
│   ├── TrayUI.cs                   # システムトレイアイコン・コンテキストメニュー
│   ├── PairingDialog.cs            # マウスペアリング設定ダイアログ
│   └── SettingsDialog.cs           # 感度・閾値設定ダイアログ
├── Configuration/                  # 設定管理レイヤー
│   ├── ConfigStore.cs              # 設定ファイルの読み書き (JSON形式)
│   └── ConfigModels.cs             # 設定データクラス (DeviceInfo, VirtualDesktopInfo等)
├── Interop/                        # Windows API相互運用
│   ├── NativeMethods.cs            # P/Invoke メソッド宣言
│   ├── NativeStructs.cs            # P/Invoke 構造体定義
│   └── NativeEnums.cs              # P/Invoke 列挙型定義
└── Properties/
    └── AssemblyInfo.cs             # アセンブリ情報
```

## 各ファイルの詳細

### 📁 ルート
- **Program.cs**
  - ApplicationContextを継承したメインクラス
  - すべてのコンポーネントの初期化と連携
  - エラーハンドリングとアプリケーションライフサイクル管理

### 📁 Core/ - コア機能
- **Events.cs**
  - `DeviceMovedEventArgs`: マウス移動イベントの引数クラス
  - デバイスID、移動量(DeltaX, DeltaY)を保持

- **VirtualDesktopController.cs**
  - 仮想デスクトップ間の切り替え制御
  - InputSimulatorを使用したキーボードショートカット送信
  - デスクトップ間でのマウス位置記憶・復元機能

- **SwitchPolicy.cs**
  - 移動量の蓄積と閾値判定
  - ヒステリシス（クールダウン）機能
  - ペアリング中の切り替え無効化制御

### 📁 Input/ - 入力処理
- **RawInputManager.cs**
  - Raw Input APIを使用したマウス入力キャプチャ
  - 複数マウスの個別識別
  - HiddenFormによるウィンドウメッセージ受信
  - デバイス移動イベントの発火

### 📁 UI/ - ユーザーインターフェース
- **TrayUI.cs**
  - システムトレイアイコンとコンテキストメニュー
  - ペアリング開始、設定画面、自動起動設定
  - 割り当て反転、バルーン通知機能

- **PairingDialog.cs**
  - マウスA/Bの識別・ペアリング設定
  - リアルタイムマウス識別UI
  - 設定の自動保存

- **SettingsDialog.cs**
  - 感度（移動閾値）とクールダウン時間の調整
  - トラックバー形式のリアルタイム設定UI

### 📁 Configuration/ - 設定管理
- **ConfigStore.cs**
  - JSON形式での設定ファイル読み書き
  - LocalApplicationData配下での設定永続化
  - デバイス-デスクトップのマッピング管理

- **ConfigModels.cs**
  - 設定データのモデルクラス群
  - JSON シリアライゼーション属性付き
  - DeviceInfo, VirtualDesktopInfo, ConfigData等

### 📁 Interop/ - Windows API相互運用
- **NativeMethods.cs**
  - P/Invokeメソッド宣言
  - RegisterRawInputDevices, GetRawInputData等
  - GetCursorPos, SetCursorPos (マウス位置制御用)

- **NativeStructs.cs**
  - Raw Input関連構造体
  - RAWINPUT, RAWMOUSE, RAWINPUTDEVICE等
  - POINT構造体 (マウス座標用)

- **NativeEnums.cs**
  - Raw Input関連列挙型
  - RawInputDeviceFlags, RawInputType等
  - RawMouseFlags, RawInputCommand

## アーキテクチャ図

```
┌─────────────────┐
│    Program.cs   │ ← エントリポイント・全体制御
└─────────┬───────┘
          │
    ┌─────▼──────────────────────────────────┐
    │              依存関係                    │
    └─┬───────┬─────────┬──────────────┬─────┘
      │       │         │              │
┌─────▼─┐ ┌──▼───┐ ┌───▼──────┐ ┌────▼──────┐
│ Input │ │ Core │ │    UI    │ │Configuration│
│Layer  │ │Layer │ │  Layer   │ │   Layer     │
└───────┘ └──────┘ └──────────┘ └─────────────┘
      │       │         │              │
      └───────┴─────────┴──────────────┴─────┐
                                             │
                                      ┌─────▼─────┐
                                      │  Interop  │
                                      │   Layer   │
                                      └───────────┘
```

## 技術スタック

- **.NET 8 (Windows Forms)**
- **Raw Input API** - 複数マウス識別
- **InputSimulator** - 仮想デスクトップ切り替え
- **Newtonsoft.Json** - 設定ファイル処理
- **P/Invoke** - Windows API呼び出し

## 主な改善点

### ✅ 保守性の向上
- 単一責任原則の適用
- 関心の分離による影響範囲の限定
- 各ファイルが明確な役割を持つ

### ✅ 可読性の向上
- 論理的なフォルダー構造
- 関連コードのグループ化
- 適切な命名規則

### ✅ 再利用性の向上
- Interop層の分離により他プロジェクトでも利用可能
- 独立性の高いコンポーネント設計

### ✅ テスト容易性
- 各コンポーネントの独立性
- 依存関係の明確化
- モックしやすいインターフェース

### ✅ 型安全性
- Nullable参照型対応
- コンパイル時警告の解消
- より堅牢なコード品質

## ビルド情報

- **ターゲット**: net8.0-windows
- **出力タイプ**: WinExe (Windows実行ファイル)
- **警告レベル**: エラー0個、警告2個（InputSimulator互換性のみ）
- **NuGetパッケージ**: InputSimulator 1.0.4, Newtonsoft.Json 13.0.3