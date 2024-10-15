#include "pch.h"
#include "GlobalEnv.h"
#include <string>
using namespace std;

int PoissonRandom(double lambda) {
	int k = 0;
	double p = 1.0;
	double L = exp(-lambda);  // L = e^(-lambda)

	do {
		k++;
		p *= static_cast<double>(rand()) / (RAND_MAX + 1.0);
	} while (p > L);

	return k - 1;
}


/////////////////////////////////////////////////////////////////////
// 파일 처리 루틴들
bool OpenFile(const CString& filename, const TCHAR* mode, FILE** fp) {

	// CString을 const wchar_t*로 변환 (유니코드 지원)
	const wchar_t* pFileName = filename.GetString();

	errno_t err = _wfopen_s(fp, pFileName, mode);
	if (err != 0 || *fp == nullptr) {
		perror("Failed to open file");
		return false;
	}
	return true;
}

void CloseFile(FILE** fp) {
	if (*fp != nullptr) {
		fclose(*fp);
		*fp = nullptr;
	}
}

ULONG WriteDataWithHeader(FILE* fp, int type, const void* data, size_t dataSize) {

	ULONG ulWritten = 0;
	ULONG ulTemp = 0;
	SAVE_TL tl = { type, static_cast<int>(dataSize) };
	ulTemp = fwrite(&tl, sizeof(tl), 1,fp);  // 먼저 데이터 타입 및 길이 정보를 쓴다
	ulWritten += ulTemp * sizeof(tl);

	ulTemp = fwrite(data, dataSize,1, fp);   // 실제 데이터 쓰기
	ulWritten += ulTemp * dataSize;

	return  ulWritten;
}

bool ReadDataWithHeader(FILE* fp, void* data, size_t expectedSize, int expectedType) {
	SAVE_TL tl;
	if (fread(&tl, 1, sizeof(tl), fp) != sizeof(tl)) {
		perror("Failed to read header");
		return false;
	}

	if (tl.type != expectedType || tl.length != expectedSize) {
		fprintf(stderr, "Data type or size mismatch\n");
		return false;
	}

	if (fread(data, 1, tl.length,fp) != tl.length) {
		perror("Failed to read data");
		return false;
	}

	return true;
}

// 확률에 따라서 0 또는 1 생성
int ZeroOrOneByProb(int probability)
{
	double randomProb = (double)rand() / RAND_MAX;
	return (randomProb <= (double)probability / 100) ? 1 : 0;
}

// 랜덤 숫자 생성 함수
int RandomBetween(int low, int high) {
	return low + rand() % (high - low + 1);
}


void clearSheet(Sheet* sheet) {
	if (sheet) {
		int lastRow = sheet->lastRow();
		int lastCol = sheet->lastCol();
		for (int i = 0; i < lastRow; i++) {
			for (int j = 0; j < lastCol; j++) {
				sheet->writeBlank(i, j, sheet->cellFormat(i, j)); // Clears content but preserves format
				sheet->setCellFormat(i, j, nullptr); // Clears format
			}
		}
	}
}


void write_project_header(Book* book, Sheet* sheet) {
	// Define header data
	const wchar_t* titles[2][16] = {
		{L"Category", L"PRJ_ID", L"기간", L"시작가능", L"끝",
		 L"발주일", L"총수익", L"경험", L"성공%", L"CF갯수",
		 L"CF1%", L"CF2%", L"CF3%", L"선금", L"중도", L"잔금"},
		{L"act갯수", L"", L"Dur", L"start", L"end",
		 L"", L"HR_H", L"HR_M", L"HR_L", L"",
		 L"mon_cf1", L"mon_cf2", L"mon_cf3", L"", L"prjType", L"actType"}
	};

	// Write data to the first and second row
	for (int row = 0; row < 2; ++row) {
		for (int col = 0; col < 16; ++col) {
			sheet->writeStr(row , col, titles[row][col]);
		}
	}

	// Define and apply styles
	Format* format = book->addFormat();
	format->setBorder(BorderStyle::BORDERSTYLE_THIN);
	format->setBorderColor(COLOR_BLACK);

	// Apply format to cells (assuming you want to apply to 1st row as an example)
	for (int col = 0; col < 16; ++col) {
		sheet->setCellFormat(0, col, format);
		sheet->setCellFormat(1, col, format);
	}
}


void write_project_body(Book* book, Sheet* sheet, PROJECT* pProject) {
	const int iWidth = 15;  // Adjusted for 0-based index, original was 16
	const int iHeight = 5;  // Adjusted for 0-based index, original was 7

	int posX, posY;

	// Adjust position for 0-based indexing
	posX = 0;  // Start at the first column (0-based)
	posY = (pProject->ID * (iHeight+1)) + 3;  // Adjust for 0-based index

	draw_outer_border(book, sheet, posY, 0, posY + iHeight-1, iWidth, BORDERSTYLE_THIN, COLOR_BLACK);

	// First row settings
	sheet->writeNum(posY, posX++, pProject->category);
	sheet->writeNum(posY, posX++, pProject->ID);
	sheet->writeNum(posY, posX++, pProject->duration);
	sheet->writeNum(posY, posX++, pProject->startAvail);
	sheet->writeNum(posY, posX++, pProject->endDate);
	sheet->writeNum(posY, posX++, pProject->orderDate);
	sheet->writeNum(posY, posX++, pProject->profit);
	sheet->writeNum(posY, posX++, pProject->experience);
	sheet->writeNum(posY, posX++, pProject->winProb);

	sheet->writeNum(posY, posX++, pProject->nCashFlows);
	sheet->writeNum(posY, posX++, pProject->cashFlows[0]);
	sheet->writeNum(posY, posX++, pProject->cashFlows[1]);
	sheet->writeNum(posY, posX++, pProject->cashFlows[2]);

	sheet->writeNum(posY, posX++, pProject->firstPay);
	sheet->writeNum(posY, posX++, pProject->secondPay);
	sheet->writeNum(posY, posX++, pProject->finalPay);

	// Second row settings
	posY++;  // Move to next row
	posX = 0;  // Reset column to start of the row
	sheet->writeNum(posY, posX++, pProject->numActivities);

	posX = 10;  // Skip to the specific column after gaps
	sheet->writeNum(posY, posX++, pProject->firstPayMonth);
	sheet->writeNum(posY, posX++, pProject->secondPayMonth);
	sheet->writeNum(posY, posX++, pProject->finalPayMonth);

	posX = 14;  // Skip to the end for final pieces of data
	sheet->writeNum(posY, posX++, pProject->projectType);
	sheet->writeNum(posY, posX++, pProject->activityPattern);

	// Activity data settings
	for (int i = 0; i < pProject->numActivities; ++i) {		

		posX = 1;  // Start from the second column (0-based)

		// Convert index to string and prepend with "Activity"
		std::wstring activityPrefix = L"Activity" + std::to_wstring(i + 1);
		sheet->writeStr(posY, posX++, activityPrefix.c_str());
		sheet->writeNum(posY, posX++, pProject->activities[i].duration);
		sheet->writeNum(posY, posX++, pProject->activities[i].startDate);
		sheet->writeNum(posY, posX++, pProject->activities[i].endDate);

		posX = 6;  // Skip two columns
		sheet->writeNum(posY, posX++, pProject->activities[i].highSkill);
		sheet->writeNum(posY, posX++, pProject->activities[i].midSkill);
		sheet->writeNum(posY, posX++, pProject->activities[i].lowSkill);

		posY++;  // Move to next row for each activity
	}
}

void draw_outer_border(Book* book, Sheet* sheet, int startRow, int startCol, int endRow, int endCol, BorderStyle borderStyle, Color borderColor) {
	if (!book || !sheet) return;

	// Apply the top border format to the top edge of the specified range
	//for (int col = startCol; col <= endCol; col++) {
	//	Format* newFormat = book->addFormat(); // Create a fresh format
	//	newFormat->setBorderTop(borderStyle);
	//	newFormat->setBorderColor(borderColor);
	//	sheet->setCellFormat(startRow, col, newFormat);
	//}

	//// Apply the left border format to the left edge of the specified range
	//for (int row = startRow; row <= endRow; row++) {
	//	Format* newFormat = book->addFormat(); // Create a fresh format
	//	newFormat->setBorderLeft(borderStyle);
	//	newFormat->setBorderColor(borderColor);
	//	sheet->setCellFormat(row, startCol, newFormat);
	//}

	//// Apply the right border format to the right edge of the specified range
	//for (int row = startRow; row <= endRow; row++) {
	//	Format* newFormat = book->addFormat(); // Create a fresh format
	//	newFormat->setBorderRight(borderStyle);
	//	newFormat->setBorderColor(borderColor);
	//	sheet->setCellFormat(row, endCol, newFormat);
	//}

	// Apply the bottom border format to the bottom edge of the specified range
	//for (int col = startCol; col <= endCol; col++) {
	//	Format* newFormat = book->addFormat(); // Create a fresh format
	//	newFormat->setBorderBottom(borderStyle);
	//	newFormat->setBorderColor(borderColor);
	//	sheet->setCellFormat(endRow, col, newFormat);
	//}
}


//void write_project_body(xlnt::worksheet& ws, PROJECT* pProject)
//{
//	const int iWidth = 16;
//	const int iHeight = 7;
//
//	int posX, posY;
//
//	
//	// 첫 번째 행 설정	  (열, 행)
//  	posX = 1; posY = pProject->ID * iHeight + 4;
//
//	draw_outer_border(ws, posY, 1, posY+ iHeight, iWidth,  xlnt::border_style::thin, xlnt::color::black());
//
//	
//
//	ws.cell(posX++, posY).value(pProject->category);
//	ws.cell(posX++, posY).value(pProject->ID);
//	ws.cell(posX++, posY).value(pProject->duration);
//	ws.cell(posX++, posY).value(pProject->startAvail);
//	ws.cell(posX++, posY).value(pProject->endDate);
//	ws.cell(posX++, posY).value(pProject->orderDate);
//	ws.cell(posX++, posY).value(pProject->profit);
//	ws.cell(posX++, posY).value(pProject->experience);
//	ws.cell(posX++, posY).value(pProject->winProb);
//	
//	ws.cell(posX++, posY).value(pProject->nCashFlows);
//	ws.cell(posX++, posY).value(pProject->cashFlows[0]);
//	ws.cell(posX++, posY).value(pProject->cashFlows[1]);
//	ws.cell(posX++, posY).value(pProject->cashFlows[2]);
//		
//	ws.cell(posX++, posY).value(pProject->firstPay);
//	ws.cell(posX++, posY).value(pProject->secondPay);
//	ws.cell(posX++, posY).value(pProject->finalPay);
//
//
//	// 두 번째 행 설정
//	posX = 1; posY ++;
//	ws.cell(posX++, posY).value(pProject->numActivities);
//
//	posX = 11;  // 빈 칸을 건너뛰기
//	ws.cell(posX++, posY).value(pProject->firstPayMonth);
//	ws.cell(posX++, posY).value(pProject->secondPayMonth);
//	ws.cell(posX++, posY).value(pProject->finalPayMonth);
//
//	posX = 15;  // 빈 칸을 건너뛰기
//	ws.cell(posX++, posY).value(pProject->projectType);
//	ws.cell(posX++, posY).value(pProject->activityPattern);
//
//	// 활동 데이터 설정
//	for (int i = 0; i < pProject->numActivities; ++i) {
//		//인덱스를 문자열로 변환하고 "Activity" 접두사 추가
//		CString strAct;
//		strAct.Format(_T("Activity%02d"), i + 1);
//		std::string str = CStringToStdString(strAct);  // CString을 std::string으로 변환
//		
//		posX = 2; // 엑셀의 2행 2열부터 적는다.
//		ws.cell(posX++, posY).value(str);
//		ws.cell(posX++, posY).value(pProject->activities[i].duration);
//		ws.cell(posX++, posY).value(pProject->activities[i].startDate);
//		ws.cell(posX++, posY).value(pProject->activities[i].endDate);
//
//		posX = 7;  // 두 열 건너뛰기
//		ws.cell(posX++, posY).value(pProject->activities[i].highSkill);
//		ws.cell(posX++, posY).value(pProject->activities[i].midSkill);
//		ws.cell(posX++, posY).value(pProject->activities[i].lowSkill);
//
//		posY++;
//	}
//
//
//	/*posX = 1; posY = pProject->ID * iHeight + 4;
//
//	draw_outer_border(ws, posY, 1, posY + iHeight, iWidth, xlnt::border_style::thin, xlnt::color::black());
//	*/
//
//	/*int printY = 4 + (pProject->ID - 1) * iHeight;
//	pXl->WriteArrayToRange(WS_NUM_PROJECT, printY, 1, (VARIANT*)projectInfo, iHeight, iWidth);
//	pXl->SetRangeBorderAround(WS_NUM_PROJECT, printY, 1, printY + iHeight - 1, iWidth + 1 - 1, 1, 2, RGB(0, 0, 0));*/
//}
