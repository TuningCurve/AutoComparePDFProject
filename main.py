# モジュール読み込み
import wx
import wx.adv
import util
import find_the_difference
import os
import sys
from pdf2image.exceptions import PDFPageCountError

if __name__ == '__main__':

    # 本プログラムがあるディレクトリを作業ディレクトリとする
    os.chdir(os.path.dirname(os.path.abspath(sys.argv[0])))

    # wxpythonの起動
    app = wx.App()

    # フレーム作成
    frame = wx.Frame(None, title="Auto Compare PDF")

    # アイコン設定（未実装）
    # icon = wx.Icon('./icon/ic.icon', wx.BITMAP_TYPE_ICO)
    # frame.SetIcon(icon)

    # パネル作成
    panel = wx.Panel(frame, -1)

    # posで指定した位置にウィジェットを配置。
    text_guide = wx.StaticText(panel, -1, 'ファイルを選択してください。', pos=(10, 10))
    text_before = wx.StaticText(panel, -1, '修正前', pos=(10, 40))
    dir_ctrl_before = wx.FilePickerCtrl(panel, -1, pos=(50, 40))
    text_after = wx.StaticText(panel, -1, '修正後', pos=(10, 80))
    dir_ctrl_after = wx.FilePickerCtrl(panel, -1, pos=(50, 80))
    btn_enter = wx.Button(panel, -1, 'エクセル比較を実行', pos=(50, 120))
    btn_enter_pdf = wx.Button(panel, -1, 'PDF比較を実行', pos=(180, 120))
    # a = wx.Bitmap('abc.png')
    # aa = wx.StaticBitmap(panel, -1, pos=(270, 80), bitmap=a)

    # エクセル比較ボタンが押されたときに呼び出される関数
    def on_click_enter(event):
        # プログレスバー
        anim = wx.Gauge(panel, -1, pos=(100, 180))
        anim.SetRange(8)
        anim.SetValue(2)
        text_progress = wx.StaticText(panel, -1, '比較処理中。。。', pos=(60, 160))
        # カレントディレクトリ取得
        cur_path = os.getcwd()
        # 入力されたファイルパスを取得
        dir_selected_before = dir_ctrl_before.GetPath()
        dir_selected_after = dir_ctrl_after.GetPath()
        if (os.path.splitext(dir_selected_before)[1] != '.xlsx') | (os.path.splitext(dir_selected_after)[1] != '.xlsx'):
            print('【エラー】選択されたファイルの拡張子が.xlsxでない')
            wx.MessageBox('エクセルを指定してください。')
            anim.Hide()
            anim.SetValue(0)
            text_progress.Hide()

        else:
            anim.SetValue(4)
            # EXCELから画像に変換し、保存
            util.excel2image(dir_selected_before, dir_selected_after, cur_path)
            anim.SetValue(6)
            # 画像の比較処理
            find_the_difference.find_the_diff('./image_file/before_01.png', './image_file/after_01.png')
            anim.SetValue(10)
            # プログレスバーを隠す
            anim.Hide()
            text_progress.Hide()
            # メッセージボックスを表示
            wx.MessageBox('比較が完了しました。')

    # PDF比較ボタンが押されたときに呼び出される関数
    def on_click_enter_pdf(event):
        # プログレスバー
        anim = wx.Gauge(panel, -1, pos=(100, 180))
        anim.SetRange(8)
        anim.SetValue(2)
        text_progress = wx.StaticText(panel, -1, '比較処理中。。。', pos=(60, 160))
        # 入力されたファイルパスを取得
        dir_selected_before = dir_ctrl_before.GetPath()
        dir_selected_after = dir_ctrl_after.GetPath()
        anim.SetValue(4)
        # PDFから画像に変換し、保存
        try:
            util.pdf2image_func(dir_selected_before, dir_selected_after)
            anim.SetValue(6)
            # 画像の比較処理
            find_the_difference.find_the_diff('./image_file/before_01.png', './image_file/after_01.png')
            anim.SetValue(10)
            # プログレスバーを隠す
            anim.Hide()
            text_progress.Hide()
            # メッセージボックスを表示
            wx.MessageBox('比較が完了しました。')

        # Image変換時エラー
        except PDFPageCountError:
            print('【エラー】PDF -> Image に変換時')
            wx.MessageBox('PDFを指定してください。')
            anim.Hide()
            anim.SetValue(0)
            text_progress.Hide()


    # 関数をボタンにバインド。
    btn_enter.Bind(wx.EVT_BUTTON, on_click_enter)
    btn_enter_pdf.Bind(wx.EVT_BUTTON, on_click_enter_pdf)

    frame.Show()
    app.MainLoop()
