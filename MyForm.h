#pragma once

namespace Project4 {
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
	private: System::Windows::Forms::DataGridView^ dataGridView1;
	private: System::Windows::Forms::Button^ button1;
	private: System::Windows::Forms::Button^ button2;
	private: System::Windows::Forms::Label^ label1;



	private: System::Windows::Forms::TextBox^ textBox1;
	private: System::Windows::Forms::Button^ button3;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ Фамилия;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ Имя;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ Возраст;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ Пол;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ Группа;


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
			this->dataGridView1 = (gcnew System::Windows::Forms::DataGridView());
			this->button1 = (gcnew System::Windows::Forms::Button());
			this->button2 = (gcnew System::Windows::Forms::Button());
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->textBox1 = (gcnew System::Windows::Forms::TextBox());
			this->button3 = (gcnew System::Windows::Forms::Button());
			this->Фамилия = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->Имя = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->Возраст = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->Пол = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->Группа = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->BeginInit();
			this->SuspendLayout();
			// 
			// dataGridView1
			// 
			this->dataGridView1->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
			this->dataGridView1->Columns->AddRange(gcnew cli::array< System::Windows::Forms::DataGridViewColumn^  >(5) {
				this->Фамилия,
					this->Имя, this->Возраст, this->Пол, this->Группа
			});
			this->dataGridView1->Location = System::Drawing::Point(12, 27);
			this->dataGridView1->Name = L"dataGridView1";
			this->dataGridView1->Size = System::Drawing::Size(544, 285);
			this->dataGridView1->TabIndex = 0;
			// 
			// button1
			// 
			this->button1->Location = System::Drawing::Point(599, 113);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(135, 23);
			this->button1->TabIndex = 1;
			this->button1->Text = L"Вывод первой строки";
			this->button1->UseVisualStyleBackColor = true;
			this->button1->Click += gcnew System::EventHandler(this, &MyForm::button1_Click);
			// 
			// button2
			// 
			this->button2->Location = System::Drawing::Point(599, 158);
			this->button2->Name = L"button2";
			this->button2->Size = System::Drawing::Size(135, 23);
			this->button2->TabIndex = 2;
			this->button2->Text = L"Вывод всех строк";
			this->button2->UseVisualStyleBackColor = true;
			this->button2->Click += gcnew System::EventHandler(this, &MyForm::button2_Click);
			// 
			// label1
			// 
			this->label1->AutoSize = true;
			this->label1->Location = System::Drawing::Point(599, 27);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(96, 13);
			this->label1->TabIndex = 3;
			this->label1->Text = L"Данные студента";
			// 
			// textBox1
			// 
			this->textBox1->Location = System::Drawing::Point(599, 70);
			this->textBox1->Name = L"textBox1";
			this->textBox1->Size = System::Drawing::Size(100, 20);
			this->textBox1->TabIndex = 4;
			// 
			// button3
			// 
			this->button3->Location = System::Drawing::Point(599, 201);
			this->button3->Name = L"button3";
			this->button3->Size = System::Drawing::Size(135, 23);
			this->button3->TabIndex = 5;
			this->button3->Text = L"Отправить в Эксель";
			this->button3->UseVisualStyleBackColor = true;
			this->button3->Click += gcnew System::EventHandler(this, &MyForm::button3_Click);
			// 
			// Фамилия
			// 
			this->Фамилия->HeaderText = L"Фамилия";
			this->Фамилия->Name = L"Фамилия";
			// 
			// Имя
			// 
			this->Имя->HeaderText = L"Имя";
			this->Имя->Name = L"Имя";
			// 
			// Возраст
			// 
			this->Возраст->HeaderText = L"Возраст";
			this->Возраст->Name = L"Возраст";
			// 
			// Пол
			// 
			this->Пол->HeaderText = L"Пол";
			this->Пол->Name = L"Пол";
			// 
			// Группа
			// 
			this->Группа->HeaderText = L"Группа";
			this->Группа->Name = L"Группа";
			// 
			// MyForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(746, 341);
			this->Controls->Add(this->button3);
			this->Controls->Add(this->textBox1);
			this->Controls->Add(this->label1);
			this->Controls->Add(this->button2);
			this->Controls->Add(this->button1);
			this->Controls->Add(this->dataGridView1);
			this->KeyPreview = true;
			this->Name = L"MyForm";
			this->Text = L"MyForm";
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->EndInit();
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion

	 
	private: System::Void Urok1ToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e) {
		 
	}
	private: System::Void button1_Click(System::Object^ sender, System::EventArgs^ e) {
		String^ filePath = "file.xlsx";
		String^ connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0 Xml;HDR=YES;'";
		OleDbConnection^ conn = gcnew OleDbConnection(connString);
		conn->Open();
		OleDbCommand^ cmd = gcnew OleDbCommand("SELECT * FROM [list$A:G]", conn);
	
		///	OleDbDataAdapter^ adapter = gcnew OleDbDataAdapter(cmd);
		OleDbDataReader^ reader = cmd->ExecuteReader();
		DataTable^ dt = gcnew DataTable();
	///	adapter->Fill(dt);
	///	dataGridView1->DataSource = dt;
		dataGridView1->Columns->Clear();
		dataGridView1->ColumnCount = 5;
		dataGridView1->Columns[0]->Name = "Фамилия";
		dataGridView1->Columns[1]->Name = "Имя";
		dataGridView1->Columns[2]->Name = "Возраст";
		dataGridView1->Columns[3]->Name = "Пол";
		dataGridView1->Columns[4]->Name = "Группа";



		if (reader->Read())
	{
	// Считывание данных из первой ячейки
	String^ cellValue1 = reader[0]->ToString();
	String^ cellValue2 = reader[1]->ToString();
	String^ cellValue3 = reader[2]->ToString();
	String^ cellValue4 = reader[3]->ToString();
	String^ cellValue5 = reader[4]->ToString();



			// Присваивание значения в textbox
	dataGridView1->Rows[0]->Cells[0]->Value = cellValue1;
	dataGridView1->Rows[0]->Cells[1]->Value = cellValue2;
	dataGridView1->Rows[0]->Cells[2]->Value = cellValue3;
	dataGridView1->Rows[0]->Cells[3]->Value = cellValue4;
	dataGridView1->Rows[0]->Cells[4]->Value = cellValue5;




	textBox1->Text = cellValue1 + " " + cellValue2;  
	}
	conn->Close();

	}
private: System::Void button2_Click(System::Object^ sender, System::EventArgs^ e) {
	String^ filePath = "file.xlsx";
	String^ connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0 Xml;HDR=YES;'";
	OleDbConnection^ conn = gcnew OleDbConnection(connString);
	conn->Open();
	OleDbCommand^ cmd = gcnew OleDbCommand("SELECT * FROM [list$A:H]", conn);
	OleDbDataAdapter^ adapter = gcnew OleDbDataAdapter(cmd);
	DataTable^ dt = gcnew DataTable();
	adapter->Fill(dt); // Заполнение DataTable данными из Excel
	dataGridView1->Columns->Clear();

	// Присвоение DataTable DataGridView
	dataGridView1->DataSource = dt;


	// Закрытие соединения с базой данных
	conn->Close();
}
private: System::Void button3_Click(System::Object^ sender, System::EventArgs^ e) {
	String^ filePath = "file.xlsx";
	String^ connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0 Xml;HDR=YES;'";
	OleDbConnection^ conn = gcnew OleDbConnection(connString);
	conn->Open();
	OleDbCommand^ cmd = gcnew OleDbCommand("CREATE TABLE[Sheet2](Фамилия VARCHAR, Имя VARCHAR, Возраст VARCHAR, Пол VARCHAR, Группа)", conn);
	 
	String^ createTableQuery = "CREATE TABLE [Sheet1] (";
	for (int i = 0; i < dataGridView1->Columns->Count; i++) {
		createTableQuery += "[" + dataGridView1->Columns[i]->HeaderText + "] VARCHAR, ";
	}
	createTableQuery = createTableQuery->Substring(0, createTableQuery->Length - 2) + ")";

	OleDbCommand^ createTableCmd = gcnew OleDbCommand(createTableQuery, conn);
	createTableCmd->ExecuteNonQuery();

	// Заполняем таблицу данными
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

