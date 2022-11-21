# pysimpleGUI関連のライブラリ
import PySimpleGUI as sg
import xlwings as xw

# 統計等のライブラリ
import numpy as np
import pandas as pd
import scipy as sp
from scipy.stats import poisson
from scipy.stats import binom
import math

# ディレクトリ用のライブラリ
import os

# 可視化ライブラリ
import matplotlib.pyplot as plt
import seaborn as sns


# ポアソン分布による金額単位サンプリングによるサンプル数算定の関数
def sample_poisson(N, pm, ke, alpha, audit_risk, internal_control='依拠しない'):
    k = np.arange(ke+1)
    pt = pm/N
    n = 1
    while True:
        mu = n*pt
        pmf_poi = poisson.cdf(k, mu)
        if pmf_poi.sum() < alpha:
            break
        n += 1
    if audit_risk == 'SR':
        n = math.ceil(n)
    if audit_risk == 'RMM-L':
        n = math.ceil(n/10*2)
    if audit_risk == 'RMM-H':
        n = math.ceil(n/2)
    if internal_control == '依拠する':
        n = math.ceil(n/3)
    return n


# GUIのテーマカラー
sg.theme('DarkBlue1')

# 各項目のレイアウト
layout = [[sg.Text('ファイル選択', font=('Arial',15)),
          sg.InputText('ファイルパス・名',key='file', font=('Arial',15)),
          sg.FilesBrowse('ファイル読込', target='file', file_types=(('Excell ファイル', '*.xlsx'),), font=('Arial',15))],
          [sg.Text('保存先の選択', font=('Arial',15)),
          sg.InputText('ファイルパス・名',key='save_file', font=('Arial',15)),
          sg.FolderBrowse('保存先を選択', target='save_file', font=('Arial',15))],
          [sg.Text('手続実施上の重要性', font=('Arial',15))],
          [sg.InputText('半角で数値を入力してください',key='pm', font=('Arial',15))],
          [sg.Text('監査リスク', font=('Arial',15))],
          [sg.Radio('RMM-L','audit_risk', font=('Arial',15)),
          sg.Radio('RMM-H','audit_risk', font=('Arial',15)),
          sg.Radio('SR','audit_risk', font=('Arial',15))],
          [sg.Text('内部統制', font=('Arial',15))],
          [sg.Radio('依拠する','internal_control', font=('Arial',15)),
          sg.Radio('依拠しない','internal_control', font=('Arial',15))],
          [sg.Text('ランダムシード', font=('Arial',15))],
          [sg.InputText('半角で数値を入力してください',key='random_state', font=('Arial',15))],
          [sg.Button('実行',key='bt', font=('Arial',15))]]


# ウィンドウ作成
window = sg.Window('202303中間監査サンプリング', layout)

# イベントループ
while True:
    event, values = window.read() #イベントの読み取り

    if event is None: # ウィンドウ閉じるとき
        break

    # エクセルファイル処理関連
    elif event == 'bt':
        file_name = values['file'] # ファイルパスを取得
        save_file_name = values['save_file'] # 保存先ファイルを指定
        amount = '金額' # 金額列のカラム名を指定
            
        if values['pm'] != '半角で数値を入力してください':
            pm = int(values['pm']) # 手続実施上の重要性
        else:
            sg.popup('半角で数値を入力してください')
        
        if values['random_state'] != '半角で数値を入力してください':
            random_state = int(values['random_state']) # ランダムシード　(サンプリングの並び替えのステータスに利用、任意の数を入力)
        else:
            sg.popup('半角で数値を入力してください')

        # 監査リスクを設定
        if values[0] == True:
            audit_risk = 'RMM-L'
        elif values[1] == True:
            audit_risk = 'RMM-H'
        elif values[2] == True:
            audit_risk = 'SR'
        else:
            sg.popup('監査リスクを設定してください')
        
        # 内部統制を設定
        if values[3] == True:
            internal_control = '依拠する'
        elif values[4] == True:
            internal_control ='依拠しない'
        else:
            sg.popup('内部統制を設定してください')
        
        # 予想虚偽表示金額（変更不要）
        ke = 0
        alpha = 0.05

        if save_file_name == 'ファイルパス・名':
            sg.popup('保存先ディレクトリを指定してください')

        if file_name != 'ファイルパス・名':
            sample_data = pd.read_excel(file_name)
            total_amount = sample_data[amount].sum()

            # サンプルサイズnの算定
            n = sample_poisson(total_amount, pm, ke, alpha, audit_risk, internal_control)

            # サンプリングシートに記載用の、パラメータ一覧
            sampling_param = pd.DataFrame([['母集団合計', total_amount],
                                        ['手続実施上の重要性', pm],
                                        ['リスク', audit_risk],
                                        ['内部統制', internal_control],
                                        ['random_state', random_state]])

            # 母集団をまずは降順に並び替える（ここで並び替えるのは、サンプル出力の安定のため安定のため）
            sample_data = sample_data.sort_values(amount, ascending=False)

            # 母集団をシャッフル
            shuffle_data = sample_data.sample(frac=1, random_state=random_state) #random_stateを使って乱数を固定化する

            # サンプリング区間の算定
            m = total_amount/n
            
            # 列の追加
            shuffle_data['cumsum'] = shuffle_data[amount].cumsum() # 積み上げ合計
            shuffle_data['group'] = shuffle_data['cumsum']//m # サンプルのグループ化

            result_data = shuffle_data.loc[shuffle_data.groupby('group')['cumsum'].idxmin()]

            file_name = '{}/サンプル.xlsx'.format(save_file_name)
            writer = pd.ExcelWriter(file_name)
            # 全レコードを'全体'シートに出力
            sample_data.to_excel(writer, sheet_name = '母集団', index=False)
            # サンプリング結果を、サンプリングシートに記載
            result_data.to_excel(writer, sheet_name = 'サンプリング結果', index=False)
            # サンプリングの情報追記
            sampling_param.to_excel(writer, sheet_name = 'サンプリングパラメータ', index=False, header=None)


            # Excelファイルを保存
            writer.save()
            # Excelファイルを閉じる
            writer.close()

            # ポップアップでメッセージ表示
            sg.popup('処理を実行しました')

        else:
            # ポップアップでメッセージ表示
            sg.popup('読み込むファイルをダウンロードしてください')

# 終了処理
window.close()