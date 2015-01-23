#pragma once

private ref class Define sealed
{
public:
	static System::String^ app_name = "板金見積";
	static System::String^ msg_quit = "ツールを終了します。";
	static System::String^ msg_save = "データを登録しますか？";
	static System::String^ msg_leave = "入力中の値がクリアされます。よろしいですか？";
	static System::String^ msg_err_export = "エラー発生：エクセル出力に失敗しました。";
	static System::String^ err_not_numeric = "数値を入力してください。";
	static System::String^ err_not_w_h_d = "サイズの値を正しく入力してください。";
};