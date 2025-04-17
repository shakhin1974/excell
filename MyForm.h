#pragma once

namespace Excell {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace System::Data::OleDb;
	using namespace Microsoft;

	/// <summary>
	/// Сводка для MyForm
	/// </summary>
	public ref class MyForm : public System::Windows::Forms::Form
	{
	public:
		MyForm(void)
		{
			InitializeComponent();
			//
			//TODO: добавьте код конструктора
			//
		}

	protected:
		/// <summary>
		/// Освободить все используемые ресурсы.
		/// </summary>
		~MyForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Button^ button1;
	protected:
	private: System::Windows::Forms::Button^ button2;
	private: System::Windows::Forms::DataGridView^ dataGridView1;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ Фамилия;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ Имя;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ Возраст;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ Пол;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ Группа;






	private:
		/// <summary>
		/// Обязательная переменная конструктора.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Требуемый метод для поддержки конструктора — не изменяйте 
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		void InitializeComponent(void)
		{
			this->button1 = (gcnew System::Windows::Forms::Button());
			this->button2 = (gcnew System::Windows::Forms::Button());
			this->dataGridView1 = (gcnew System::Windows::Forms::DataGridView());
			this->Фамилия = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->Имя = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->Возраст = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->Пол = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->Группа = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->BeginInit();
			this->SuspendLayout();
			// 
			// button1
			// 
			this->button1->Location = System::Drawing::Point(44, 13);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(75, 23);
			this->button1->TabIndex = 0;
			this->button1->Text = L"Create DB";
			this->button1->UseVisualStyleBackColor = true;
			this->button1->Click += gcnew System::EventHandler(this, &MyForm::button1_Click);
			// 
			// button2
			// 
			this->button2->Location = System::Drawing::Point(629, 12);
			this->button2->Name = L"button2";
			this->button2->Size = System::Drawing::Size(75, 23);
			this->button2->TabIndex = 1;
			this->button2->Text = L"Insert DB";
			this->button2->UseVisualStyleBackColor = true;
			this->button2->Click += gcnew System::EventHandler(this, &MyForm::button2_Click);
			// 
			// dataGridView1
			// 
			this->dataGridView1->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
			this->dataGridView1->Columns->AddRange(gcnew cli::array< System::Windows::Forms::DataGridViewColumn^  >(5) {
				this->Фамилия,
					this->Имя, this->Возраст, this->Пол, this->Группа
			});
			this->dataGridView1->Location = System::Drawing::Point(44, 109);
			this->dataGridView1->Name = L"dataGridView1";
			this->dataGridView1->Size = System::Drawing::Size(592, 215);
			this->dataGridView1->TabIndex = 2;
			// 
			// Фамилия
			// 
			this->Фамилия->HeaderText = L"Фамилия";
			this->Фамилия->Name = L"Фамилия";
			this->Фамилия->Width = 150;
			// 
			// Имя
			// 
			this->Имя->HeaderText = L"Имя";
			this->Имя->Name = L"Имя";
			this->Имя->Width = 150;
			// 
			// Возраст
			// 
			this->Возраст->HeaderText = L"Возраст";
			this->Возраст->Name = L"Возраст";
			this->Возраст->Width = 50;
			// 
			// Пол
			// 
			this->Пол->HeaderText = L"Пол";
			this->Пол->Name = L"Пол";
			this->Пол->Width = 50;
			// 
			// Группа
			// 
			this->Группа->HeaderText = L"Группа";
			this->Группа->Name = L"Группа";
			this->Группа->Width = 150;
			// 
			// MyForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(1337, 563);
			this->Controls->Add(this->dataGridView1);
			this->Controls->Add(this->button2);
			this->Controls->Add(this->button1);
			this->Name = L"MyForm";
			this->Text = L"MyForm";
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->EndInit();
			this->ResumeLayout(false);

		}

bool SheetExists(OleDbConnection^ conn, String^ sheetName) {
			try {
				OleDbCommand^ cmd = gcnew OleDbCommand(
					"SELECT TOP 1 * FROM [" + sheetName + "$]",
					conn
				);
				cmd->ExecuteReader()->Close(); // Попытка прочитать лист
				return true;
			}
			catch (OleDbException^) {
				return false; // Лист не существует
			}
		}

#pragma endregion
	private: System::Void button1_Click(System::Object^ sender, System::EventArgs^ e) {
		String^ connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=file.xls;Extended Properties=\"Excel 8.0;HDR=YES;\"";
		OleDbConnection^ conn = gcnew OleDbConnection(connString);
		conn->Open();

		String^ sheetName = "Sheet1";
		if (!SheetExists(conn, sheetName)) {
			OleDbCommand^ cmd = gcnew OleDbCommand(
				"CREATE TABLE [" + sheetName + "$] ([Фамилия] TEXT(255), [Имя] TEXT(255), [Возраст] TEXT(50))",
				conn
			);
			cmd->ExecuteNonQuery();
			MessageBox::Show("Таблица создана!");
		}
		else {
			MessageBox::Show("Таблица уже существует!");
		}

		OleDbCommand^ cmd = gcnew OleDbCommand("CREATE TABLE[Sheet2](Фамилия VARCHAR, Имя VARCHAR, Возраст VARCHAR, Пол VARCHAR, Группа VARCHAR)", conn);
		
	
		
		
		
		conn->Close();


	}
private: System::Void button2_Click(System::Object^ sender, System::EventArgs^ e) {
	String^ connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=file.xls;Extended Properties=\"Excel 8.0;HDR=YES;\"";
	OleDbConnection^ conn = gcnew OleDbConnection(connString);
	conn->Open();

	for (int i = 0; i < dataGridView1->Rows->Count; i++) {
		if (dataGridView1->Rows[i]->IsNewRow) continue; // Пропускаем новую строку

		String^ insertQuery = "INSERT INTO [Sheet1] VALUES (";
		for (int j = 0; j < dataGridView1->Columns->Count; j++) {
			insertQuery += "'" + dataGridView1->Rows[i]->Cells[j]->Value->ToString() + "', ";
		}
		insertQuery = insertQuery->Substring(0, insertQuery->Length - 2) + ")";

		OleDbCommand^ insertCmd = gcnew OleDbCommand(insertQuery, conn);
		insertCmd->ExecuteNonQuery();

	}
	

	conn->Close();
}
};
}
