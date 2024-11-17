import pandas as pd
import numpy as np
from typing import List, Tuple, Optional
import random
import os
import sys
import time
import logging
from datetime import datetime

class DataAnalysisError(Exception):
    """カスタムエラークラス"""
    pass

def setup_logging():
    """ログ設定"""
    log_dir = 'logs'
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file = os.path.join(log_dir, f'analysis_{timestamp}.log')
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler(sys.stdout)
        ]
    )

def print_progress_bar(iteration: int, total: int, prefix: str = '', suffix: str = '', length: int = 50, fill: str = '█'):
    """プログレスバーの表示"""
    percent = ("{0:.1f}").format(100 * (iteration / float(total)))
    filled_length = int(length * iteration // total)
    bar = fill * filled_length + '-' * (length - filled_length)
    print(f'\r{prefix} |{bar}| {percent}% {suffix}', end='\r')
    if iteration == total:
        print()

def check_excel_file_validity(file_path: str) -> Tuple[bool, str]:
    """
    Excelファイルの妥当性チェック
    
    Returns:
        Tuple[bool, str]: (妥当性, エラーメッセージ)
    """
    try:
        if not os.path.exists(file_path):
            return False, f"ファイルが存在しません: {file_path}"
        
        if not file_path.endswith(('.xlsx', '.xls')):
            return False, f"不適切なファイル形式です。Excelファイル(.xlsx, .xls)である必要があります: {file_path}"
        
        # ファイルが開けるかチェック
        xl = pd.ExcelFile(file_path)
        if not xl.sheet_names:
            return False, f"有効なシートが存在しません: {file_path}"
            
        return True, ""
        
    except Exception as e:
        return False, f"ファイルチェック中にエラーが発生しました: {str(e)}"

def get_excel_files_from_directory(directory: str) -> List[str]:
    """
    指定されたディレクトリからExcelファイルを取得
    """
    try:
        if not os.path.exists(directory):
            raise DataAnalysisError(
                f"'data'ディレクトリが見つかりません。\n"
                f"以下の手順で準備してください：\n"
                f"1. 'analyze_data.py'と同じフォルダに'data'フォルダを作成\n"
                f"2. 'data'フォルダに分析対象のExcelファイル('.xlsx' or '.xls')を配置してください。"
            )

        excel_files = [f for f in os.listdir(directory) 
                      if f.endswith(('.xlsx', '.xls'))]
        
        if not excel_files:
            raise DataAnalysisError(
                f"'data'フォルダにExcelファイルが見つかりません。\n"
                f"分析対象のExcelファイル(.xlsx or .xls)を配置してください。"
            )
        
        # ファイルの妥当性チェック
        full_paths = [os.path.join(directory, f) for f in excel_files]
        for file_path in full_paths:
            is_valid, error_msg = check_excel_file_validity(file_path)
            if not is_valid:
                raise DataAnalysisError(f"ファイルチェックエラー: {error_msg}")
        
        return sorted(full_paths)  # ファイル名でソート
        
    except DataAnalysisError:
        raise
    except Exception as e:
        raise DataAnalysisError(f"予期せぬエラーが発生しました: {str(e)}")

def get_sheet_data(file_path: str) -> List[pd.DataFrame]:
    """
    Excelファイルから有効なシートのデータを取得
    """
    try:
        xl = pd.ExcelFile(file_path)
        valid_sheets = []
        invalid_sheets = []
        
        # FutureWarningを抑制
        import warnings
        warnings.filterwarnings('ignore', category=FutureWarning)
        
        for sheet_name in xl.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            if df.empty or df.isnull().all().all():
                invalid_sheets.append((sheet_name, "空のシート"))
                continue
                
            df = df.iloc[1:, 1:]  # 1行目と1列目を削除
            
            # applymap の代わりに apply を使用
            non_numeric = df.apply(lambda col: col.apply(
                lambda x: not pd.api.types.is_numeric_dtype(type(x)) and not pd.isnull(x)
            ))
            
            if non_numeric.any().any():
                non_numeric_pos = [(i+2, j+2) for i, j in zip(*non_numeric.values.nonzero())]
                invalid_sheets.append((sheet_name, f"数値以外のデータが含まれています: {non_numeric_pos}"))
                continue
                
            df = df.apply(pd.to_numeric, errors='coerce').fillna(0)
            valid_sheets.append(df)
        
        if not valid_sheets:
            error_details = "\n".join(f"- シート'{name}': {reason}" for name, reason in invalid_sheets)
            raise DataAnalysisError(
                f"有効なデータを含むシートが見つかりませんでした: {file_path}\n"
                f"エラー詳細:\n{error_details}"
            )
        
        return valid_sheets
        
    except Exception as e:
        raise DataAnalysisError(f"ファイル読み込み中にエラーが発生しました ({file_path}): {str(e)}")

def validate_iteration_input(input_value: str) -> Optional[int]:
    """
    ランダム化検定の試行回数の入力値を検証
    """
    try:
        n_iterations = int(input_value)
        if n_iterations <= 0:
            print("エラー: 試行回数は正の整数である必要があります")
            return None
        if n_iterations > 1000000:
            print("警告: 試行回数が大きすぎます。処理に時間がかかる可能性があります。")
            confirmation = input("続行しますか？ (y/n): ").lower()
            if confirmation != 'y':
                return None
        return n_iterations
    except ValueError:
        print("エラー: 有効な数値を入力してください")
        return None

def process_data_frames(dfs: List[pd.DataFrame]) -> pd.DataFrame:
    """データフレームのリストを処理し、合計を返す"""
    if not dfs:
        raise DataAnalysisError("処理対象のデータが空です")
        
    try:
        total_df = dfs[0].copy()
        for df in dfs[1:]:
            if df.shape != total_df.shape:
                raise DataAnalysisError(
                    f"シート間でデータサイズが一致しません\n"
                    f"期待サイズ: {total_df.shape}\n"
                    f"実際のサイズ: {df.shape}"
                )
            total_df += df
            
        return total_df
        
    except DataAnalysisError:
        raise
    except Exception as e:
        raise DataAnalysisError(f"データ処理中にエラーが発生しました: {str(e)}")

def create_cross_tabulation(df: pd.DataFrame) -> pd.DataFrame:
    """クロス集計表を作成"""
    try:
        row_totals = df.sum(axis=1)
        col_totals = df.sum(axis=0)
        grand_total = df.values.sum()

        df_copy = df.copy()
        df_copy['Row Total'] = row_totals
        col_totals_with_grand = pd.concat([col_totals, pd.Series({'Row Total': grand_total})])
        df_copy.loc['Column Total'] = col_totals_with_grand

        return df_copy
        
    except Exception as e:
        raise DataAnalysisError(f"クロス集計表の作成中にエラーが発生しました: {str(e)}")

def save_cross_tabulation(df: pd.DataFrame, base_filename: str, output_dir: str = 'output') -> None:
    """
    クロス集計表をExcelファイルとして保存
    """
    try:
        # 出力ディレクトリの作成（存在しない場合）
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # 現在の日時を取得してファイル名に追加
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f"{base_filename}_crosstab_{timestamp}.xlsx"
        output_path = os.path.join(output_dir, output_filename)
        
        # Excelファイルとして保存
        df.to_excel(output_path)
        
    except Exception as e:
        print(f"\nクロス集計表の保存中にエラーが発生しました: {e}")

def calculate_expected_frequencies(observed_df: pd.DataFrame) -> pd.DataFrame:
    """期待度数を計算"""
    try:
        total = float(observed_df.loc['Column Total', 'Row Total'])
        row_totals = observed_df['Row Total'].drop('Column Total').astype(float)
        col_totals = observed_df.loc['Column Total'].drop('Row Total').astype(float)

        expected_df = observed_df.copy().astype(float)
        
        for i in row_totals.index:
            for j in col_totals.index:
                expected_df.loc[i, j] = (row_totals[i] * col_totals[j]) / total

        return expected_df
        
    except Exception as e:
        raise DataAnalysisError(f"期待度数の計算中にエラーが発生しました: {str(e)}")

def calculate_chi_square(observed_df: pd.DataFrame, expected_df: pd.DataFrame) -> Tuple[float, pd.Series]:
    """カイ二乗値を計算"""
    try:
        observed = observed_df.drop('Column Total').drop('Row Total', axis=1)
        expected = expected_df.drop('Column Total').drop('Row Total', axis=1)
        
        # ゼロ除算を回避
        with np.errstate(divide='ignore', invalid='ignore'):
            chi_square_elements = (observed - expected) ** 2 / expected
            chi_square_elements = chi_square_elements.replace([np.inf, -np.inf], 0).fillna(0)
            row_chi_squares = chi_square_elements.sum(axis=1)
        
        return row_chi_squares.sum(), row_chi_squares
        
    except Exception as e:
        raise DataAnalysisError(f"カイ二乗値の計算中にエラーが発生しました: {str(e)}")

def analyze_data(data_frames: List[pd.DataFrame]) -> float:
    """データフレームのリストを分析し、カイ二乗値を返す"""
    try:
        total_df = process_data_frames(data_frames)
        cross_tab = create_cross_tabulation(total_df)
        expected_freq = calculate_expected_frequencies(cross_tab)
        chi_square, _ = calculate_chi_square(cross_tab, expected_freq)
        return chi_square
        
    except Exception as e:
        raise DataAnalysisError(f"データ分析中にエラーが発生しました: {str(e)}")

def main():
    """メイン実行関数"""
    try:
        # FutureWarningを抑制
        import warnings
        warnings.filterwarnings('ignore', category=FutureWarning)
        
        print("\n=== データ分析プログラム ===")
        print("\n前提条件:")
        print("1. スクリプトと同じフォルダに'data'フォルダが必要です")
        print("2. 'data'フォルダに分析対象のExcelファイルを配置してください")
        print("3. 各Excelファイルには、分析対象のデータが含まれている必要があります")
        print("\n処理を開始します...\n")
        
        # データディレクトリの取得と作成
        script_dir = os.path.dirname(os.path.abspath(__file__))
        data_dir = os.path.join(script_dir, 'data')
        output_dir = os.path.join(script_dir, 'output')
        
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)
            raise DataAnalysisError(
                f"'data'ディレクトリを作成しました: {data_dir}\n"
                f"以下の手順で準備してください:\n"
                f"1. 作成された'data'フォルダに分析対象のExcelファイルを配置\n"
                f"2. プログラムを再実行"
            )
        
        # Excelファイルの取得
        excel_files = get_excel_files_from_directory(data_dir)
        print("分析対象ファイル:")
        for i, file in enumerate(excel_files, 1):
            print(f"{i}: {os.path.basename(file)}")
        
        # データの読み込みと検証
        print("\nファイルを読み込んでいます...")
        all_sheets_list = []
        sheet_counts = []
        
        for file in excel_files:
            sheets = get_sheet_data(file)
            all_sheets_list.append(sheets)
            sheet_counts.append(len(sheets))
        
        total_sheets = sum(sheet_counts)
        all_sheets = [sheet for sheets in all_sheets_list for sheet in sheets]
        
        print(f"\n合計有効シート数: {total_sheets}")
        
        # 元のデータの分析とクロス集計表の保存
        print("\nデータを分析しています...")
        
        # クロス集計表と期待度数の作成と保存
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            
        for i, sheets in enumerate(all_sheets_list):
            total_df = process_data_frames(sheets)
            cross_tab = create_cross_tabulation(total_df)
            expected_freq = calculate_expected_frequencies(cross_tab)
            
            base_filename = os.path.splitext(os.path.basename(excel_files[i]))[0]
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            
            # クロス集計表の保存
            output_filename = f"{base_filename}_crosstab_{timestamp}.xlsx"
            output_path = os.path.join(output_dir, output_filename)
            cross_tab.to_excel(output_path)
            
            # 期待度数の保存
            expected_output_filename = f"{base_filename}_expected_{timestamp}.xlsx"
            expected_output_path = os.path.join(output_dir, expected_output_filename)
            expected_freq.to_excel(expected_output_path)
        
        # 元のカイ二乗値の計算
        original_chi_square = sum(analyze_data(sheets) for sheets in all_sheets_list)
        print(f"元のデータの全体カイ二乗値: {original_chi_square:.6f}")
        
        # ランダム化テストのパラメータ入力と検証
        while True:
            try:
                print("\n--- ランダム化検定の設定 ---")
                print("推奨試行回数: 10000")
                print("注意: 大きな値を設定すると処理時間が長くなります")
                input_value = input("試行回数を入力してください: ")
                n_iterations = validate_iteration_input(input_value)
                if n_iterations is not None:
                    break
            except KeyboardInterrupt:
                print("\n\n処理を中断します。")
                return
            except Exception as e:
                print(f"\nエラーが発生しました: {e}")
                print("もう一度入力してください。")
        
        # ランダム化テストの実行
        print("\nランダム化検定を開始します...")
        count_greater_equal = 0
        
        start_time = time.time()
        
        try:
            for i in range(n_iterations):
                # シートをシャッフル
                random.shuffle(all_sheets)
                
                # グループに分割
                random_groups = []
                start_idx = 0
                for count in sheet_counts:
                    end_idx = start_idx + count
                    random_groups.append(all_sheets[start_idx:end_idx])
                    start_idx = end_idx
                
                # ランダム化したデータの分析
                random_chi_square = sum(analyze_data(group) for group in random_groups)
                
                ##### オリジナルのカイ二乗値が大きい場合にカウントする #####
                ##### もし、オリジナルのカイ二乗値"以上"をカウントしたい場合は、不等号を">="に変更する #####
                if original_chi_square > random_chi_square:
                    count_greater_equal += 1
                
                # 進捗の表示
                if n_iterations >= 10 and (i % (n_iterations // 10) == 0 or i == n_iterations - 1):
                    print_progress_bar(i + 1, n_iterations, prefix='進捗:', suffix='完了', length=50)
            
            # 結果の表示
            proportion = count_greater_equal / n_iterations
            total_time = time.time() - start_time
            
            print("\n=== 分析結果 ===")
            print(f"処理時間: {total_time:.1f}秒")
            print(f"試行回数: {n_iterations}")
            print(f"元のカイ二乗値: {original_chi_square:.6f}")
            print(f"元のカイ二乗値の方が大きい回数: {count_greater_equal}")
            print(f"元のカイ二乗値の方が大きい割合: {proportion:.4f}")
            
        except KeyboardInterrupt:
            print("\n\n処理を中断しました。")
        except Exception as e:
            print(f"\n予期せぬエラーが発生しました: {e}")
        finally:
            print("\nプログラムを終了します。")

    except KeyboardInterrupt:
        print("\n\nプログラムを中断しました。")
    except DataAnalysisError as e:
        print(f"\nエラー: {e}")
    except Exception as e:
        print(f"\n予期せぬエラーが発生しました: {e}")

if __name__ == "__main__":
    main()
