#pragma once

namespace MetalPriceTallier {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// FormMain の概要
	/// </summary>
	public ref class FormMain : public System::Windows::Forms::Form
	{
	public:
		FormMain(void)
		{
			InitializeComponent();
			define = gcnew Define();
			check = gcnew CheckInput();
			InitInput();
			//init objects
			lbBankin->SelectedIndex = 0;
			bResist->Enabled = false;
			bExcelExport->Enabled = false;
		}
	protected:
		/// <summary>
		/// 使用中のリソースをすべてクリーンアップします。
		/// </summary>
		~FormMain()
		{
			if (components) {
				delete components;
			}
		}
	private:
		//internal classes
		Define^ define;
		CheckInput^ check;
		//objects on Form
		System::Windows::Forms::ListBox^  lbBankin;
		System::Windows::Forms::GroupBox^  gbInput;
		System::Windows::Forms::PictureBox^  pbBox;
		System::Windows::Forms::ComboBox^  cbType;
		System::Windows::Forms::Label^  lType;
		System::Windows::Forms::Label^  lPrice;
		System::Windows::Forms::Label^  lCharge;
		System::Windows::Forms::Label^  lUnitPrice;
		System::Windows::Forms::Label^  lDimension;
		System::Windows::Forms::Label^  lD;
		System::Windows::Forms::TextBox^  tbD;
		System::Windows::Forms::Label^  lH;
		System::Windows::Forms::TextBox^  tbH;
		System::Windows::Forms::Label^  lW;
		System::Windows::Forms::TextBox^  tbW;
		System::Windows::Forms::ComboBox^  cbQuantity;
		System::Windows::Forms::Label^  lQuantitiy;
		System::Windows::Forms::GroupBox^  gbTotal;
		System::Windows::Forms::Label^  lBankinTotal;
		System::Windows::Forms::Label^  label20;
		System::Windows::Forms::Label^  lTitle;
		System::Windows::Forms::Label^  lVersion;
		System::Windows::Forms::PictureBox^  pbLogo;
		System::Windows::Forms::ListBox^  lbWeight;
		System::Windows::Forms::ListBox^  lbKakohi;
		System::Windows::Forms::Button^  bQuit;
		System::Windows::Forms::Label^  lWeight;
		System::Windows::Forms::Label^  lKakohi;
		System::Windows::Forms::ComboBox^  cbCharge;
		System::Windows::Forms::ComboBox^  cbUnitPrice;
		System::Windows::Forms::TextBox^  tbPrice;
		System::Windows::Forms::TextBox^  tbDimension;
		System::Windows::Forms::TextBox^  tbTotal;
		System::Windows::Forms::TextBox^  tbBankinTotal;
		System::Windows::Forms::Button^  bClear;
		System::Windows::Forms::Button^  bResist;
		System::Windows::Forms::Button^  bExcelExport;
		System::Windows::Forms::CheckBox^  cbKadai;
		System::Windows::Forms::TextBox^  tbColorEdit;
		System::Windows::Forms::Label^  lHM;
		System::Windows::Forms::Label^  lDM;
		System::Windows::Forms::Label^  lWM;	
	private: System::Windows::Forms::Button^  bDelete;
			 /// <summary>
		/// 必要なデザイナー変数です。
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// デザイナー サポートに必要なメソッドです。このメソッドの内容を
		/// コード エディターで変更しないでください。
		/// </summary>
		void InitializeComponent(void)
		{
			System::ComponentModel::ComponentResourceManager^  resources = (gcnew System::ComponentModel::ComponentResourceManager(FormMain::typeid));
			this->lbBankin = (gcnew System::Windows::Forms::ListBox());
			this->gbInput = (gcnew System::Windows::Forms::GroupBox());
			this->lHM = (gcnew System::Windows::Forms::Label());
			this->lDM = (gcnew System::Windows::Forms::Label());
			this->lWM = (gcnew System::Windows::Forms::Label());
			this->cbKadai = (gcnew System::Windows::Forms::CheckBox());
			this->tbPrice = (gcnew System::Windows::Forms::TextBox());
			this->tbDimension = (gcnew System::Windows::Forms::TextBox());
			this->cbCharge = (gcnew System::Windows::Forms::ComboBox());
			this->cbUnitPrice = (gcnew System::Windows::Forms::ComboBox());
			this->bClear = (gcnew System::Windows::Forms::Button());
			this->bResist = (gcnew System::Windows::Forms::Button());
			this->pbBox = (gcnew System::Windows::Forms::PictureBox());
			this->cbType = (gcnew System::Windows::Forms::ComboBox());
			this->lType = (gcnew System::Windows::Forms::Label());
			this->lPrice = (gcnew System::Windows::Forms::Label());
			this->lCharge = (gcnew System::Windows::Forms::Label());
			this->lUnitPrice = (gcnew System::Windows::Forms::Label());
			this->lDimension = (gcnew System::Windows::Forms::Label());
			this->lD = (gcnew System::Windows::Forms::Label());
			this->tbD = (gcnew System::Windows::Forms::TextBox());
			this->lH = (gcnew System::Windows::Forms::Label());
			this->tbH = (gcnew System::Windows::Forms::TextBox());
			this->lW = (gcnew System::Windows::Forms::Label());
			this->tbW = (gcnew System::Windows::Forms::TextBox());
			this->cbQuantity = (gcnew System::Windows::Forms::ComboBox());
			this->lQuantitiy = (gcnew System::Windows::Forms::Label());
			this->gbTotal = (gcnew System::Windows::Forms::GroupBox());
			this->tbColorEdit = (gcnew System::Windows::Forms::TextBox());
			this->bExcelExport = (gcnew System::Windows::Forms::Button());
			this->tbTotal = (gcnew System::Windows::Forms::TextBox());
			this->tbBankinTotal = (gcnew System::Windows::Forms::TextBox());
			this->lBankinTotal = (gcnew System::Windows::Forms::Label());
			this->label20 = (gcnew System::Windows::Forms::Label());
			this->lTitle = (gcnew System::Windows::Forms::Label());
			this->lVersion = (gcnew System::Windows::Forms::Label());
			this->pbLogo = (gcnew System::Windows::Forms::PictureBox());
			this->lbWeight = (gcnew System::Windows::Forms::ListBox());
			this->lbKakohi = (gcnew System::Windows::Forms::ListBox());
			this->bQuit = (gcnew System::Windows::Forms::Button());
			this->lWeight = (gcnew System::Windows::Forms::Label());
			this->lKakohi = (gcnew System::Windows::Forms::Label());
			this->bDelete = (gcnew System::Windows::Forms::Button());
			this->gbInput->SuspendLayout();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^  >(this->pbBox))->BeginInit();
			this->gbTotal->SuspendLayout();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^  >(this->pbLogo))->BeginInit();
			this->SuspendLayout();
			// 
			// lbBankin
			// 
			this->lbBankin->Font = (gcnew System::Drawing::Font(L"メイリオ", 8.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point, 
				static_cast<System::Byte>(128)));
			this->lbBankin->FormattingEnabled = true;
			this->lbBankin->ItemHeight = 17;
			this->lbBankin->Items->AddRange(gcnew cli::array< System::Object^  >(8) {L"1)", L"2)", L"3)", L"4)", L"5)", L"6)", L"7)", 
				L"8)"});
			this->lbBankin->Location = System::Drawing::Point(15, 56);
			this->lbBankin->Name = L"lbBankin";
			this->lbBankin->Size = System::Drawing::Size(499, 140);
			this->lbBankin->TabIndex = 0;
			this->lbBankin->SelectedIndexChanged += gcnew System::EventHandler(this, &FormMain::lbBankin_SelectedIndexChanged);
			// 
			// gbInput
			// 
			this->gbInput->Controls->Add(this->lHM);
			this->gbInput->Controls->Add(this->lDM);
			this->gbInput->Controls->Add(this->lWM);
			this->gbInput->Controls->Add(this->cbKadai);
			this->gbInput->Controls->Add(this->tbPrice);
			this->gbInput->Controls->Add(this->tbDimension);
			this->gbInput->Controls->Add(this->cbCharge);
			this->gbInput->Controls->Add(this->cbUnitPrice);
			this->gbInput->Controls->Add(this->bClear);
			this->gbInput->Controls->Add(this->bResist);
			this->gbInput->Controls->Add(this->pbBox);
			this->gbInput->Controls->Add(this->cbType);
			this->gbInput->Controls->Add(this->lType);
			this->gbInput->Controls->Add(this->lPrice);
			this->gbInput->Controls->Add(this->lCharge);
			this->gbInput->Controls->Add(this->lUnitPrice);
			this->gbInput->Controls->Add(this->lDimension);
			this->gbInput->Controls->Add(this->lD);
			this->gbInput->Controls->Add(this->tbD);
			this->gbInput->Controls->Add(this->lH);
			this->gbInput->Controls->Add(this->tbH);
			this->gbInput->Controls->Add(this->lW);
			this->gbInput->Controls->Add(this->tbW);
			this->gbInput->Controls->Add(this->cbQuantity);
			this->gbInput->Controls->Add(this->lQuantitiy);
			this->gbInput->Font = (gcnew System::Drawing::Font(L"メイリオ", 9, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point, 
				static_cast<System::Byte>(128)));
			this->gbInput->Location = System::Drawing::Point(15, 214);
			this->gbInput->Name = L"gbInput";
			this->gbInput->Size = System::Drawing::Size(426, 347);
			this->gbInput->TabIndex = 2;
			this->gbInput->TabStop = false;
			this->gbInput->Text = L"板金設定";
			// 
			// lHM
			// 
			this->lHM->AutoSize = true;
			this->lHM->Location = System::Drawing::Point(165, 141);
			this->lHM->Name = L"lHM";
			this->lHM->Size = System::Drawing::Size(20, 18);
			this->lHM->TabIndex = 31;
			this->lHM->Text = L"m";
			// 
			// lDM
			// 
			this->lDM->AutoSize = true;
			this->lDM->Location = System::Drawing::Point(165, 173);
			this->lDM->Name = L"lDM";
			this->lDM->Size = System::Drawing::Size(20, 18);
			this->lDM->TabIndex = 30;
			this->lDM->Text = L"m";
			// 
			// lWM
			// 
			this->lWM->AutoSize = true;
			this->lWM->Location = System::Drawing::Point(165, 106);
			this->lWM->Name = L"lWM";
			this->lWM->Size = System::Drawing::Size(20, 18);
			this->lWM->TabIndex = 29;
			this->lWM->Text = L"m";
			// 
			// cbKadai
			// 
			this->cbKadai->AutoSize = true;
			this->cbKadai->Location = System::Drawing::Point(211, 65);
			this->cbKadai->Name = L"cbKadai";
			this->cbKadai->Size = System::Drawing::Size(51, 22);
			this->cbKadai->TabIndex = 2;
			this->cbKadai->Text = L"架台";
			this->cbKadai->UseVisualStyleBackColor = true;
			this->cbKadai->CheckedChanged += gcnew System::EventHandler(this, &FormMain::cbKadai_CheckedChanged);
			// 
			// tbPrice
			// 
			this->tbPrice->Location = System::Drawing::Point(89, 296);
			this->tbPrice->Name = L"tbPrice";
			this->tbPrice->ReadOnly = true;
			this->tbPrice->Size = System::Drawing::Size(96, 25);
			this->tbPrice->TabIndex = 28;
			this->tbPrice->TabStop = false;
			this->tbPrice->Text = L"0";
			this->tbPrice->TextAlign = System::Windows::Forms::HorizontalAlignment::Right;
			this->tbPrice->TextChanged += gcnew System::EventHandler(this, &FormMain::tbPrice_TextChanged);
			// 
			// tbDimension
			// 
			this->tbDimension->Location = System::Drawing::Point(89, 201);
			this->tbDimension->Name = L"tbDimension";
			this->tbDimension->ReadOnly = true;
			this->tbDimension->Size = System::Drawing::Size(66, 25);
			this->tbDimension->TabIndex = 6;
			this->tbDimension->TabStop = false;
			this->tbDimension->TextAlign = System::Windows::Forms::HorizontalAlignment::Right;
			this->tbDimension->TextChanged += gcnew System::EventHandler(this, &FormMain::tbDimension_TextChanged);
			// 
			// cbCharge
			// 
			this->cbCharge->FormatString = L"N0";
			this->cbCharge->FormattingEnabled = true;
			this->cbCharge->Items->AddRange(gcnew cli::array< System::Object^  >(6) {L"500", L"600", L"700", L"800", L"900", L"1000"});
			this->cbCharge->Location = System::Drawing::Point(89, 264);
			this->cbCharge->Name = L"cbCharge";
			this->cbCharge->Size = System::Drawing::Size(66, 26);
			this->cbCharge->TabIndex = 7;
			this->cbCharge->SelectedIndexChanged += gcnew System::EventHandler(this, &FormMain::cbCharge_SelectedIndexChanged);
			this->cbCharge->TextChanged += gcnew System::EventHandler(this, &FormMain::cbCharge_TextChanged);
			// 
			// cbUnitPrice
			// 
			this->cbUnitPrice->FormatString = L"N3";
			this->cbUnitPrice->FormattingEnabled = true;
			this->cbUnitPrice->Items->AddRange(gcnew cli::array< System::Object^  >(11) {L"15000", L"16000", L"17000", L"18000", L"19000", 
				L"20000", L"21000", L"22000", L"23000", L"24000", L"25000"});
			this->cbUnitPrice->Location = System::Drawing::Point(89, 232);
			this->cbUnitPrice->Name = L"cbUnitPrice";
			this->cbUnitPrice->Size = System::Drawing::Size(66, 26);
			this->cbUnitPrice->TabIndex = 6;
			this->cbUnitPrice->SelectedIndexChanged += gcnew System::EventHandler(this, &FormMain::cbUnitPrice_SelectedIndexChanged);
			this->cbUnitPrice->TextChanged += gcnew System::EventHandler(this, &FormMain::cbUnitPrice_TextChanged);
			// 
			// bClear
			// 
			this->bClear->Font = (gcnew System::Drawing::Font(L"MS UI Gothic", 9, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point, 
				static_cast<System::Byte>(128)));
			this->bClear->Location = System::Drawing::Point(298, 219);
			this->bClear->Name = L"bClear";
			this->bClear->Size = System::Drawing::Size(89, 37);
			this->bClear->TabIndex = 8;
			this->bClear->Text = L"クリア";
			this->bClear->UseVisualStyleBackColor = true;
			this->bClear->Click += gcnew System::EventHandler(this, &FormMain::bClear_Click);
			// 
			// bResist
			// 
			this->bResist->Font = (gcnew System::Drawing::Font(L"MS UI Gothic", 9, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point, 
				static_cast<System::Byte>(128)));
			this->bResist->Location = System::Drawing::Point(298, 273);
			this->bResist->Name = L"bResist";
			this->bResist->Size = System::Drawing::Size(89, 50);
			this->bResist->TabIndex = 9;
			this->bResist->Text = L"確 定";
			this->bResist->UseVisualStyleBackColor = true;
			this->bResist->Click += gcnew System::EventHandler(this, &FormMain::bResist_Click);
			// 
			// pbBox
			// 
			this->pbBox->Image = (cli::safe_cast<System::Drawing::Image^  >(resources->GetObject(L"pbBox.Image")));
			this->pbBox->Location = System::Drawing::Point(285, 25);
			this->pbBox->Name = L"pbBox";
			this->pbBox->Size = System::Drawing::Size(135, 170);
			this->pbBox->TabIndex = 22;
			this->pbBox->TabStop = false;
			// 
			// cbType
			// 
			this->cbType->DropDownStyle = System::Windows::Forms::ComboBoxStyle::DropDownList;
			this->cbType->FormattingEnabled = true;
			this->cbType->Items->AddRange(gcnew cli::array< System::Object^  >(3) {L"制御盤", L"操作盤", L"中継盤"});
			this->cbType->Location = System::Drawing::Point(88, 63);
			this->cbType->Name = L"cbType";
			this->cbType->Size = System::Drawing::Size(97, 26);
			this->cbType->TabIndex = 1;
			this->cbType->SelectedIndexChanged += gcnew System::EventHandler(this, &FormMain::cbType_SelectedIndexChanged);
			// 
			// lType
			// 
			this->lType->AutoSize = true;
			this->lType->Location = System::Drawing::Point(30, 66);
			this->lType->Name = L"lType";
			this->lType->Size = System::Drawing::Size(32, 18);
			this->lType->TabIndex = 3;
			this->lType->Text = L"種別";
			// 
			// lPrice
			// 
			this->lPrice->AutoSize = true;
			this->lPrice->Location = System::Drawing::Point(30, 305);
			this->lPrice->Name = L"lPrice";
			this->lPrice->Size = System::Drawing::Size(32, 18);
			this->lPrice->TabIndex = 15;
			this->lPrice->Text = L"金額";
			// 
			// lCharge
			// 
			this->lCharge->AutoSize = true;
			this->lCharge->Location = System::Drawing::Point(30, 271);
			this->lCharge->Name = L"lCharge";
			this->lCharge->Size = System::Drawing::Size(44, 18);
			this->lCharge->TabIndex = 14;
			this->lCharge->Text = L"その他";
			// 
			// lUnitPrice
			// 
			this->lUnitPrice->AutoSize = true;
			this->lUnitPrice->Location = System::Drawing::Point(30, 238);
			this->lUnitPrice->Name = L"lUnitPrice";
			this->lUnitPrice->Size = System::Drawing::Size(32, 18);
			this->lUnitPrice->TabIndex = 12;
			this->lUnitPrice->Text = L"単価";
			// 
			// lDimension
			// 
			this->lDimension->AutoSize = true;
			this->lDimension->Location = System::Drawing::Point(30, 206);
			this->lDimension->Name = L"lDimension";
			this->lDimension->Size = System::Drawing::Size(32, 18);
			this->lDimension->TabIndex = 9;
			this->lDimension->Text = L"総㎡";
			// 
			// lD
			// 
			this->lD->AutoSize = true;
			this->lD->Location = System::Drawing::Point(30, 173);
			this->lD->Name = L"lD";
			this->lD->Size = System::Drawing::Size(17, 18);
			this->lD->TabIndex = 8;
			this->lD->Text = L"D";
			// 
			// tbD
			// 
			this->tbD->Location = System::Drawing::Point(88, 170);
			this->tbD->Name = L"tbD";
			this->tbD->Size = System::Drawing::Size(66, 25);
			this->tbD->TabIndex = 5;
			this->tbD->TextAlign = System::Windows::Forms::HorizontalAlignment::Right;
			this->tbD->TextChanged += gcnew System::EventHandler(this, &FormMain::tbD_TextChanged);
			// 
			// lH
			// 
			this->lH->AutoSize = true;
			this->lH->Location = System::Drawing::Point(30, 141);
			this->lH->Name = L"lH";
			this->lH->Size = System::Drawing::Size(17, 18);
			this->lH->TabIndex = 6;
			this->lH->Text = L"H";
			// 
			// tbH
			// 
			this->tbH->Location = System::Drawing::Point(88, 138);
			this->tbH->Name = L"tbH";
			this->tbH->Size = System::Drawing::Size(66, 25);
			this->tbH->TabIndex = 4;
			this->tbH->TextAlign = System::Windows::Forms::HorizontalAlignment::Right;
			this->tbH->TextChanged += gcnew System::EventHandler(this, &FormMain::tbH_TextChanged);
			// 
			// lW
			// 
			this->lW->AutoSize = true;
			this->lW->Location = System::Drawing::Point(30, 106);
			this->lW->Name = L"lW";
			this->lW->Size = System::Drawing::Size(20, 18);
			this->lW->TabIndex = 4;
			this->lW->Text = L"W";
			// 
			// tbW
			// 
			this->tbW->Location = System::Drawing::Point(88, 103);
			this->tbW->Name = L"tbW";
			this->tbW->Size = System::Drawing::Size(66, 25);
			this->tbW->TabIndex = 3;
			this->tbW->TextAlign = System::Windows::Forms::HorizontalAlignment::Right;
			this->tbW->TextChanged += gcnew System::EventHandler(this, &FormMain::tbW_TextChanged);
			// 
			// cbQuantity
			// 
			this->cbQuantity->AutoCompleteCustomSource->AddRange(gcnew cli::array< System::String^  >(5) {L"1", L"2", L"3", L"4", L"5"});
			this->cbQuantity->DropDownStyle = System::Windows::Forms::ComboBoxStyle::DropDownList;
			this->cbQuantity->FormattingEnabled = true;
			this->cbQuantity->Items->AddRange(gcnew cli::array< System::Object^  >(5) {L"1", L"2", L"3", L"4", L"5"});
			this->cbQuantity->Location = System::Drawing::Point(88, 27);
			this->cbQuantity->Name = L"cbQuantity";
			this->cbQuantity->Size = System::Drawing::Size(67, 26);
			this->cbQuantity->TabIndex = 0;
			this->cbQuantity->SelectedIndexChanged += gcnew System::EventHandler(this, &FormMain::cbQuantity_SelectedIndexChanged);
			// 
			// lQuantitiy
			// 
			this->lQuantitiy->AutoSize = true;
			this->lQuantitiy->Location = System::Drawing::Point(30, 30);
			this->lQuantitiy->Name = L"lQuantitiy";
			this->lQuantitiy->Size = System::Drawing::Size(32, 18);
			this->lQuantitiy->TabIndex = 1;
			this->lQuantitiy->Text = L"数量";
			// 
			// gbTotal
			// 
			this->gbTotal->Controls->Add(this->tbColorEdit);
			this->gbTotal->Controls->Add(this->bExcelExport);
			this->gbTotal->Controls->Add(this->tbTotal);
			this->gbTotal->Controls->Add(this->tbBankinTotal);
			this->gbTotal->Controls->Add(this->lBankinTotal);
			this->gbTotal->Controls->Add(this->label20);
			this->gbTotal->Font = (gcnew System::Drawing::Font(L"メイリオ", 9, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point, 
				static_cast<System::Byte>(128)));
			this->gbTotal->Location = System::Drawing::Point(541, 42);
			this->gbTotal->Name = L"gbTotal";
			this->gbTotal->Size = System::Drawing::Size(197, 198);
			this->gbTotal->TabIndex = 17;
			this->gbTotal->TabStop = false;
			this->gbTotal->Text = L"合計";
			// 
			// tbColorEdit
			// 
			this->tbColorEdit->Location = System::Drawing::Point(85, 57);
			this->tbColorEdit->Name = L"tbColorEdit";
			this->tbColorEdit->Size = System::Drawing::Size(66, 25);
			this->tbColorEdit->TabIndex = 1;
			this->tbColorEdit->Text = L"0";
			this->tbColorEdit->TextAlign = System::Windows::Forms::HorizontalAlignment::Right;
			this->tbColorEdit->TextChanged += gcnew System::EventHandler(this, &FormMain::tbColorEdit_TextChanged);
			// 
			// bExcelExport
			// 
			this->bExcelExport->Font = (gcnew System::Drawing::Font(L"MS UI Gothic", 9, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point, 
				static_cast<System::Byte>(128)));
			this->bExcelExport->Image = (cli::safe_cast<System::Drawing::Image^  >(resources->GetObject(L"bExcelExport.Image")));
			this->bExcelExport->Location = System::Drawing::Point(62, 129);
			this->bExcelExport->Name = L"bExcelExport";
			this->bExcelExport->Size = System::Drawing::Size(89, 50);
			this->bExcelExport->TabIndex = 3;
			this->bExcelExport->UseVisualStyleBackColor = true;
			this->bExcelExport->Click += gcnew System::EventHandler(this, &FormMain::bExcelExport_Click);
			// 
			// tbTotal
			// 
			this->tbTotal->Location = System::Drawing::Point(85, 88);
			this->tbTotal->Name = L"tbTotal";
			this->tbTotal->ReadOnly = true;
			this->tbTotal->Size = System::Drawing::Size(66, 25);
			this->tbTotal->TabIndex = 2;
			this->tbTotal->TabStop = false;
			this->tbTotal->Text = L"0";
			this->tbTotal->TextAlign = System::Windows::Forms::HorizontalAlignment::Right;
			this->tbTotal->TextChanged += gcnew System::EventHandler(this, &FormMain::tbTotal_TextChanged);
			// 
			// tbBankinTotal
			// 
			this->tbBankinTotal->Location = System::Drawing::Point(85, 24);
			this->tbBankinTotal->Name = L"tbBankinTotal";
			this->tbBankinTotal->ReadOnly = true;
			this->tbBankinTotal->Size = System::Drawing::Size(66, 25);
			this->tbBankinTotal->TabIndex = 0;
			this->tbBankinTotal->TabStop = false;
			this->tbBankinTotal->Text = L"0";
			this->tbBankinTotal->TextAlign = System::Windows::Forms::HorizontalAlignment::Right;
			// 
			// lBankinTotal
			// 
			this->lBankinTotal->AutoSize = true;
			this->lBankinTotal->Location = System::Drawing::Point(26, 30);
			this->lBankinTotal->Name = L"lBankinTotal";
			this->lBankinTotal->Size = System::Drawing::Size(32, 18);
			this->lBankinTotal->TabIndex = 9;
			this->lBankinTotal->Text = L"板金";
			// 
			// label20
			// 
			this->label20->AutoSize = true;
			this->label20->Location = System::Drawing::Point(26, 64);
			this->label20->Name = L"label20";
			this->label20->Size = System::Drawing::Size(32, 18);
			this->label20->TabIndex = 1;
			this->label20->Text = L"調色";
			// 
			// lTitle
			// 
			this->lTitle->AutoSize = true;
			this->lTitle->Font = (gcnew System::Drawing::Font(L"メイリオ", 12, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point, 
				static_cast<System::Byte>(128)));
			this->lTitle->Location = System::Drawing::Point(92, 19);
			this->lTitle->Name = L"lTitle";
			this->lTitle->Size = System::Drawing::Size(122, 24);
			this->lTitle->TabIndex = 0;
			this->lTitle->Text = L"板金見積ツール";
			// 
			// lVersion
			// 
			this->lVersion->AutoSize = true;
			this->lVersion->Font = (gcnew System::Drawing::Font(L"メイリオ", 9.75F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point, 
				static_cast<System::Byte>(128)));
			this->lVersion->Location = System::Drawing::Point(237, 22);
			this->lVersion->Name = L"lVersion";
			this->lVersion->Size = System::Drawing::Size(57, 20);
			this->lVersion->TabIndex = 19;
			this->lVersion->Text = L"Ver.1.0";
			// 
			// pbLogo
			// 
			this->pbLogo->Image = (cli::safe_cast<System::Drawing::Image^  >(resources->GetObject(L"pbLogo.Image")));
			this->pbLogo->Location = System::Drawing::Point(31, 9);
			this->pbLogo->Name = L"pbLogo";
			this->pbLogo->Size = System::Drawing::Size(36, 35);
			this->pbLogo->SizeMode = System::Windows::Forms::PictureBoxSizeMode::StretchImage;
			this->pbLogo->TabIndex = 20;
			this->pbLogo->TabStop = false;
			// 
			// lbWeight
			// 
			this->lbWeight->Font = (gcnew System::Drawing::Font(L"メイリオ", 9, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point, 
				static_cast<System::Byte>(128)));
			this->lbWeight->FormattingEnabled = true;
			this->lbWeight->ItemHeight = 18;
			this->lbWeight->Items->AddRange(gcnew cli::array< System::Object^  >(10) {L"サイズ", L"20ｘ20ｘ3　　0.885", L"25ｘ25ｘ3　　1.120", L"30ｘ30ｘ3　　1.360", 
				L"30ｘ30ｘ5　　2.160", L"40ｘ40ｘ3　　1.830", L"45ｘ45ｘ4　　2.740", L"45ｘ45ｘ5　　3.380", L"50ｘ50ｘ4　　3.060", L"50ｘ50ｘ6　　4.430"});
			this->lbWeight->Location = System::Drawing::Point(468, 296);
			this->lbWeight->Name = L"lbWeight";
			this->lbWeight->Size = System::Drawing::Size(138, 184);
			this->lbWeight->TabIndex = 22;
			this->lbWeight->TabStop = false;
			// 
			// lbKakohi
			// 
			this->lbKakohi->Font = (gcnew System::Drawing::Font(L"メイリオ", 9, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point, 
				static_cast<System::Byte>(128)));
			this->lbKakohi->FormattingEnabled = true;
			this->lbKakohi->ItemHeight = 18;
			this->lbKakohi->Items->AddRange(gcnew cli::array< System::Object^  >(2) {L"　　 円/台", L"700円 or 800円"});
			this->lbKakohi->Location = System::Drawing::Point(631, 296);
			this->lbKakohi->Name = L"lbKakohi";
			this->lbKakohi->SelectionMode = System::Windows::Forms::SelectionMode::None;
			this->lbKakohi->Size = System::Drawing::Size(108, 58);
			this->lbKakohi->TabIndex = 23;
			this->lbKakohi->TabStop = false;
			// 
			// bQuit
			// 
			this->bQuit->Location = System::Drawing::Point(649, 494);
			this->bQuit->Name = L"bQuit";
			this->bQuit->Size = System::Drawing::Size(89, 37);
			this->bQuit->TabIndex = 25;
			this->bQuit->Text = L"閉じる";
			this->bQuit->UseVisualStyleBackColor = true;
			this->bQuit->Click += gcnew System::EventHandler(this, &FormMain::bQuit_Click);
			// 
			// lWeight
			// 
			this->lWeight->AutoSize = true;
			this->lWeight->Font = (gcnew System::Drawing::Font(L"メイリオ", 9.75F));
			this->lWeight->Location = System::Drawing::Point(477, 260);
			this->lWeight->Name = L"lWeight";
			this->lWeight->Size = System::Drawing::Size(87, 20);
			this->lWeight->TabIndex = 26;
			this->lWeight->Text = L"架台重量一覧";
			// 
			// lKakohi
			// 
			this->lKakohi->AutoSize = true;
			this->lKakohi->Font = (gcnew System::Drawing::Font(L"メイリオ", 9.75F));
			this->lKakohi->Location = System::Drawing::Point(646, 260);
			this->lKakohi->Name = L"lKakohi";
			this->lKakohi->Size = System::Drawing::Size(48, 20);
			this->lKakohi->TabIndex = 27;
			this->lKakohi->Text = L"加工費";
			// 
			// bDelete
			// 
			this->bDelete->Location = System::Drawing::Point(442, 22);
			this->bDelete->Name = L"bDelete";
			this->bDelete->Size = System::Drawing::Size(72, 29);
			this->bDelete->TabIndex = 28;
			this->bDelete->Text = L"行削除";
			this->bDelete->UseVisualStyleBackColor = true;
			this->bDelete->Click += gcnew System::EventHandler(this, &FormMain::bDelete_Click);
			// 
			// FormMain
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 12);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(761, 579);
			this->Controls->Add(this->bDelete);
			this->Controls->Add(this->lKakohi);
			this->Controls->Add(this->lWeight);
			this->Controls->Add(this->bQuit);
			this->Controls->Add(this->lbKakohi);
			this->Controls->Add(this->lbWeight);
			this->Controls->Add(this->pbLogo);
			this->Controls->Add(this->lVersion);
			this->Controls->Add(this->lTitle);
			this->Controls->Add(this->gbTotal);
			this->Controls->Add(this->gbInput);
			this->Controls->Add(this->lbBankin);
			this->Name = L"FormMain";
			this->gbInput->ResumeLayout(false);
			this->gbInput->PerformLayout();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^  >(this->pbBox))->EndInit();
			this->gbTotal->ResumeLayout(false);
			this->gbTotal->PerformLayout();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^  >(this->pbLogo))->EndInit();
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
		//button click
		System::Void bExcelExport_Click(System::Object^  sender, System::EventArgs^  e)
		{
			//Call C# com object
			array<System::String^> ^items = gcnew array<System::String^>(8);			
			for(int i = 0; i < 8; i++) {
				items[i] = lbBankin->Items[i]->ToString();				
			}
			//Marunage
			if (!ExcelOutputLib::ExcelOutput::Output(items)) {
				MessageBox::Show(define->msg_err_export, define->app_name, MessageBoxButtons::OK, MessageBoxIcon::Error);
			}
		}
		System::Void bResist_Click(System::Object^  sender, System::EventArgs^  e)
		{
			UpdateBankin();
			is_edit = false;
			lbBankin->SelectedIndex = pre_index;
			InitInput();
		}
		System::Void bClear_Click(System::Object^  sender, System::EventArgs^  e)
		{
			InitInput();
		}
		System::Void bQuit_Click(System::Object^  sender, System::EventArgs^  e)
		{
			if (MessageBox::Show(define->msg_quit, define->app_name, 
				MessageBoxButtons::OKCancel, MessageBoxIcon::Information) == Windows::Forms::DialogResult::OK) {
				Application::Exit();
			}
		}
		System::Void bDelete_Click(System::Object^  sender, System::EventArgs^  e)
		{
			if (lbBankin->SelectedIndex != -1)
			{
				array<System::String^> ^items = gcnew array<System::String^>(8);
				//Get present items
				for(int i = 0; i < 8; i++) {
					if (lbBankin->SelectedIndex != i) {
						items[i] = lbBankin->Items[i]->ToString();
					} else {
						//clear selected item
						items[i] = Convert::ToString(i + 1) + ")";
					}
				}
				lbBankin->Items->Clear();
				for(int i = 0; i < 8; i++) {
					lbBankin->Items->Add(items[i]);
				}
				tbBankinTotal->Text = GetBankinTotal();
				SetTotalPrice();
			}
		}
		//value changed
		System::Void cbQuantity_SelectedIndexChanged(System::Object^  sender, System::EventArgs^  e)
		{
			is_edit = true;
		}
		System::Void tbW_TextChanged(System::Object^  sender, System::EventArgs^  e)
		{
			is_edit = true;
			if (CheckWHD(tbW->Text)) {
				SetDimension();
			} else {
				tbDimension->Text = "";
			}
		}
		System::Void tbH_TextChanged(System::Object^  sender, System::EventArgs^  e)
		{
			is_edit = true;
			if (CheckWHD(tbH->Text)) {
				SetDimension();
			} else {
				tbDimension->Text = "";
			}
		}
		System::Void tbD_TextChanged(System::Object^  sender, System::EventArgs^  e)
		{
			is_edit = true;
			if (CheckWHD(tbD->Text)) {
				SetDimension();
			} else {
				tbDimension->Text = "";
			}
		}
		System::Void cbKadai_CheckedChanged(System::Object^  sender, System::EventArgs^  e)
		{
			is_edit = true;
			SetKadai();
		}
		System::Void tbDimension_TextChanged(System::Object^  sender, System::EventArgs^  e)
		{
			SetPrice();
		}
		System::Void cbUnitPrice_SelectedIndexChanged(System::Object^  sender, System::EventArgs^  e)
		{
			is_edit = true;
			SetPrice();
		}
		System::Void cbUnitPrice_TextChanged(System::Object^  sender, System::EventArgs^  e)
		{
			is_edit = true;
			if (CheckNumeric(cbUnitPrice->Text)) {
				SetPrice();
			} else {
				tbPrice->Text = "";
			}
		}
		System::Void cbCharge_TextChanged(System::Object^  sender, System::EventArgs^  e)
		{
			is_edit = true;
			if ((!cbKadai->Checked && cbCharge->Text == "") || CheckNumeric(cbCharge->Text)) {
				SetPrice();
			} else {
				tbPrice->Text = "";
			}
		}
		System::Void tbPrice_TextChanged(System::Object^  sender, System::EventArgs^  e)
		{
			if (tbPrice->Text != "" && cbType->Text != "")
			{
				bResist->Enabled = true;
			} else {
				bResist->Enabled = false;
			}
		}
		System::Void tbColorEdit_TextChanged(System::Object^  sender, System::EventArgs^  e)
		{
			if (CheckNumeric(tbColorEdit->Text)) {
				SetTotalPrice();
			} else {
				tbTotal->Text = "";
			}
		}
		System::Void tbTotal_TextChanged(System::Object^  sender, System::EventArgs^  e)
		{
			if (tbTotal->Text != "")
			{
				bExcelExport->Enabled = true;
			} else {
				bExcelExport->Enabled = false;
			}
		}
		//selected index changed
		System::Void cbType_SelectedIndexChanged(System::Object^  sender, System::EventArgs^  e)
		{
			if (tbPrice->Text != "" && cbType->Text != "")
			{
				bResist->Enabled = true;
			} else {
				bResist->Enabled = false;
			}
		}
		System::Void cbCharge_SelectedIndexChanged(System::Object^  sender, System::EventArgs^  e)
		{
			is_edit = true;
			SetPrice();
		}
		System::Void lbBankin_SelectedIndexChanged(System::Object^  sender, System::EventArgs^  e)
		{
			array<System::String^> ^phrazes = lbBankin->SelectedItems[0]->ToString()->Split(' ');
			if (phrazes->Length > 1) {
				if (!is_edit || MessageBox::Show(define->msg_leave, define->app_name,
					MessageBoxButtons::OKCancel, MessageBoxIcon::Question) == Windows::Forms::DialogResult::OK) {
					InitInput();
					//Set listbox item into textboxes
					cbQuantity->Text = phrazes[1]->Substring(phrazes[1]->IndexOf(':') + 1);
					cbType->Text = phrazes[2]->Substring(phrazes[2]->IndexOf(':') + 1);
					cbKadai->Checked = (phrazes[3] == "架台"? true: false);
					//架台、盤名				
					array<System::String^> ^sizes = phrazes[4]->Split('x');
					tbW->Text = sizes[0];
					tbH->Text = sizes[1];
					tbD->Text = sizes[2];
					cbUnitPrice->Text = phrazes[6]->Substring(phrazes[6]->IndexOf(':') + 1);
					cbCharge->Text = phrazes[7]->Substring(phrazes[7]->IndexOf(':') + 1);
					pre_index = lbBankin->SelectedIndex;
				} else {
					is_edit = false;
					lbBankin->SelectedIndex = pre_index;
					is_edit = true;
				}
			} else {
				pre_index = lbBankin->SelectedIndex;
			}
		}
		//private Methods
		bool CheckNumeric(String^ input) {
			bool is_check_ok = false;
			if (input != "") {
				if (!check->IsNumeric(input)) {
					MessageBox::Show(define->err_not_numeric, define->app_name, 
					MessageBoxButtons::OK, MessageBoxIcon::Error);
				} else {
					is_check_ok = true;
				}
			}
			return is_check_ok;
		}
		bool CheckWHD(String^ input) {
			bool is_check_ok = false;
			if (input != "") {
				if (!check->IsWHD(input)) {
					MessageBox::Show(define->err_not_w_h_d, define->app_name, 
					MessageBoxButtons::OK, MessageBoxIcon::Error);
				} else {
					is_check_ok = true;
				}
			}
			return is_check_ok;
		}
		void SetDimension() {
			//Validate again in case not justifyed yet
			if (check->IsWHD(tbW->Text) && check->IsWHD(tbH->Text) && check->IsWHD(tbD->Text)) {
				Decimal dimension;
				if (cbKadai->Checked) {
					//Kadai dimension
					dimension = (Convert::ToDecimal(tbW->Text) +
						Convert::ToDecimal(tbH->Text) +
						Convert::ToDecimal(tbD->Text)) * 4;
				} else {
					//Ban dimension
					dimension = (Convert::ToDecimal(tbW->Text) * Convert::ToDecimal(tbH->Text) + 
						Convert::ToDecimal(tbH->Text) * Convert::ToDecimal(tbD->Text) + 
						Convert::ToDecimal(tbD->Text) * Convert::ToDecimal(tbW->Text)) * 2;  
				}
				tbDimension->Text = Convert::ToString(dimension);
			} else {
				tbDimension->Text = "";
			}
		}
		void SetKadai() 
		{
			//Set labels
			lDimension->Text = this->cbKadai->Checked? "全長m": "総㎡";
			lUnitPrice->Text = this->cbKadai->Checked? "重さ/m": "単価";
			lCharge->Text = this->cbKadai->Checked? "加工費": "その他";
			//Reset combobox items
			cbUnitPrice->SelectedIndex = -1;
			cbCharge->SelectedIndex = -1;
			cbUnitPrice->Items->Clear();
			if (this->cbKadai->Checked) {
				cbUnitPrice->Items->Add("0.885");
				cbUnitPrice->Items->Add("1.120");
				cbUnitPrice->Items->Add("1.360");
				cbUnitPrice->Items->Add("2.160");
				cbUnitPrice->Items->Add("1.830");
				cbUnitPrice->Items->Add("2.950");
				cbUnitPrice->Items->Add("2.740");
				cbUnitPrice->Items->Add("3.380");
				cbUnitPrice->Items->Add("3.060");
				cbUnitPrice->Items->Add("4.430");
			} else {
				cbUnitPrice->Items->Add("15000");
				cbUnitPrice->Items->Add("16000");
				cbUnitPrice->Items->Add("17000");
				cbUnitPrice->Items->Add("18000");
				cbUnitPrice->Items->Add("19000");
				cbUnitPrice->Items->Add("20000");
				cbUnitPrice->Items->Add("21000");
				cbUnitPrice->Items->Add("22000");
				cbUnitPrice->Items->Add("23000");
				cbUnitPrice->Items->Add("24000");
				cbUnitPrice->Items->Add("25000");
			}
			//Reset price
			SetDimension();
			tbPrice->Text = "";
			SetPrice();
		}
		void SetPrice()
		{
			//Validate again in case not justifyed yet
			if (tbDimension->Text != "" && check->IsNumeric(cbUnitPrice->Text) && 
				((!cbKadai->Checked && cbCharge->Text == "") || check->IsNumeric(cbCharge->Text))) {
				Decimal price;
				if (cbKadai->Checked) {
					//Kadai price
					price = Math::Ceiling(
						Convert::ToDecimal(tbDimension->Text) *
						Convert::ToDecimal(cbUnitPrice->Text) *
						Convert::ToDecimal(cbCharge->Text) / 100) * 100;
				} else {
					//Ban price
					price = Math::Ceiling( 
						Convert::ToDecimal(tbDimension->Text) *
						Convert::ToDecimal(cbUnitPrice->Text) +
						(cbCharge->Text == ""? 0 : Convert::ToDecimal(cbCharge->Text)) / 100) * 100;
				}
				tbPrice->Text = Convert::ToString(price);
			} else {
				tbPrice->Text = "";
			}
		}
		void UpdateBankin()
		{
			array<System::String^> ^items = gcnew array<System::String^>(8);
			//Get present items
			for(int i = 0; i < 8; i++) {
				if (lbBankin->SelectedIndex != i) {
					items[i] = lbBankin->Items[i]->ToString();
				} else {
					//Set selected item via textboxes & labels
					items[i] = Convert::ToString(i + 1) + ") "								//No
							+ "数量:" + cbQuantity->Text + " "								//Qty
							+ cbType->Text + " " + (cbKadai->Checked? "架台":"") + " "		//type
							+ tbW->Text + "x" + tbH->Text + "x" + tbD->Text + " "			//size
							+ lDimension->Text + ":" + tbDimension->Text + " "				//dimension
							+ lUnitPrice->Text + ":" + cbUnitPrice->Text + " "				//unit price
							+ lCharge->Text + ":" + cbCharge->Text + " "					//charge
							+ lPrice->Text + ":" + tbPrice->Text ;							//price
				}
			}			
			lbBankin->Items->Clear();
			for(int i = 0; i < 8; i++) {
				lbBankin->Items->Add(items[i]);
			}
			tbBankinTotal->Text = GetBankinTotal();
			SetTotalPrice();
		}
		void SetTotalPrice()
		{
			if (tbBankinTotal->Text != "0" && check->IsNumeric(tbColorEdit->Text)) {
				tbTotal->Text = Convert::ToString(
					Convert::ToDecimal(tbBankinTotal->Text) +
					Convert::ToDecimal(tbColorEdit->Text));
			} else {
				tbTotal->Text = "";
			}
		}
		void InitInput()
		{
			cbQuantity->SelectedIndex = 0;
			//Init is_edit flag
			is_edit = false;

			cbType->SelectedIndex = -1;
			tbW->Text = "";
			tbH->Text = "";
			tbD->Text = "";
			tbDimension->Text = "";
			cbUnitPrice->SelectedIndex = -1;
			cbUnitPrice->Text = "";
			cbCharge->SelectedIndex = -1;
			cbCharge->Text = "";
			tbPrice->Text = "";
			is_edit = false;
		}
		System::String^ GetBankinTotal()
		{
			Decimal total;
			for(int i = 0; i < 8; i++) {
				array<System::String^> ^phrazes = lbBankin->Items[i]->ToString()->Split(' ');
				if (phrazes->Length > 1) {					
					Decimal qty = Convert::ToDecimal(
						lbBankin->Items[i]->ToString()->Substring(
						lbBankin->Items[i]->ToString()->IndexOf(':') + 1, 1));
					Decimal unit_price = Convert::ToDecimal(
						lbBankin->Items[i]->ToString()->Substring(
						lbBankin->Items[i]->ToString()->LastIndexOf(':') + 1));
					total += qty * unit_price;
				}
			}
			return Convert::ToString(total);
		}
		property bool is_edit;
		property int pre_index;
	};
}

