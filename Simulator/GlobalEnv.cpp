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

	int posY = 4; //엑셀의 5번 행 부터기록
	
	// Write data to the first and second row
	for (int row = posY; row < posY+2; ++row) {
		for (int col = 0; col < 16; ++col) {
			sheet->writeStr(row , col, titles[row- posY][col]);
		}
	}

	Format* format = book->addFormat();
	format->setBorder(BorderStyle::BORDERSTYLE_THIN);
	format->setBorderColor(COLOR_BLACK);

	// Apply format to cells (assuming you want to apply to 1st row as an example)
	for (int col = 0; col < 16; ++col) {
		sheet->setCellFormat(posY,	col, format);
		sheet->setCellFormat(posY+1, col, format);
	}
}


void write_project_body(Book* book, Sheet* sheet, PROJECT* pProject) {
	const int iWidth = 15;  // Adjusted for 0-based index, original was 16
	const int iHeight = 5;  // Adjusted for 0-based index, original was 7

	int posX, posY;

	// Adjust position for 0-based indexing
	posX = 0;  // Start at the first column (0-based)
	posY = (pProject->ID * (iHeight+1)) + 7;  // Adjust for 0-based index

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


void read_project_body(Book* book, Sheet* sheet, PROJECT* pProject, int projectIndex) {
	const int iWidth = 15;
	const int iHeight = 5;

	int posX, posY;

	posX = 0;
	posY = (projectIndex * (iHeight + 1)) + 7;

	// First row readings
	pProject->category = sheet->readNum(posY, posX++);
	pProject->ID = sheet->readNum(posY, posX++);
	pProject->duration = sheet->readNum(posY, posX++);
	pProject->startAvail = sheet->readNum(posY, posX++);
	pProject->endDate = sheet->readNum(posY, posX++);
	pProject->orderDate = sheet->readNum(posY, posX++);
	pProject->profit = sheet->readNum(posY, posX++);
	pProject->experience = sheet->readNum(posY, posX++);
	pProject->winProb = sheet->readNum(posY, posX++);

	pProject->nCashFlows = sheet->readNum(posY, posX++);
	pProject->cashFlows[0] = sheet->readNum(posY, posX++);
	pProject->cashFlows[1] = sheet->readNum(posY, posX++);
	pProject->cashFlows[2] = sheet->readNum(posY, posX++);

	pProject->firstPay = sheet->readNum(posY, posX++);
	pProject->secondPay = sheet->readNum(posY, posX++);
	pProject->finalPay = sheet->readNum(posY, posX++);

	// Second row readings
	posY++;  //주의!!!
	posX = 0;
	pProject->numActivities = sheet->readNum(posY, posX++);

	posX = 10;
	pProject->firstPayMonth = sheet->readNum(posY, posX++);
	pProject->secondPayMonth = sheet->readNum(posY, posX++);
	pProject->finalPayMonth = sheet->readNum(posY, posX++);

	posX = 14;
	pProject->projectType = sheet->readNum(posY, posX++);
	pProject->activityPattern = sheet->readNum(posY, posX++);

	// Activity data readings
	for (int i = 0; i < pProject->numActivities; ++i) {		
		posX = 1;

		// Skip activity name
		posX++;

		pProject->activities[i].duration = sheet->readNum(posY, posX++);
		pProject->activities[i].startDate = sheet->readNum(posY, posX++);
		pProject->activities[i].endDate = sheet->readNum(posY, posX++);

		posX = 6;
		pProject->activities[i].highSkill = sheet->readNum(posY, posX++);
		pProject->activities[i].midSkill = sheet->readNum(posY, posX++);
		pProject->activities[i].lowSkill = sheet->readNum(posY, posX++);

		posY++;
	}
}


void draw_outer_border(Book* book, Sheet* sheet, int startRow, int startCol, int endRow, int endCol, BorderStyle borderStyle, Color borderColor)
{
	if (!book || !sheet || startRow > endRow || startCol > endCol)
		return;

	return; // libxl 에서 회신이 올때까지는 동작 시키지 말자

	// 각 면과 모서리의 테두리를 위한 포맷 생성
	Format* topFormat = book->addFormat();
	Format* bottomFormat = book->addFormat();
	Format* leftFormat = book->addFormat();
	Format* rightFormat = book->addFormat();
	Format* topLeftFormat = book->addFormat();
	Format* topRightFormat = book->addFormat();
	Format* bottomLeftFormat = book->addFormat();
	Format* bottomRightFormat = book->addFormat();

	// 내부 테두리를 지우기 위한 포맷 생성
	Format* clearFormat = book->addFormat();
	clearFormat->setBorder(BORDERSTYLE_NONE);

	// 테두리 스타일과 색상 설정
	topFormat->setBorderTop(borderStyle);
	bottomFormat->setBorderBottom(borderStyle);
	leftFormat->setBorderLeft(borderStyle);
	rightFormat->setBorderRight(borderStyle);
	topFormat->setBorderTopColor(borderColor);
	bottomFormat->setBorderBottomColor(borderColor);
	leftFormat->setBorderLeftColor(borderColor);
	rightFormat->setBorderRightColor(borderColor);

	// 모서리 포맷 설정
	topLeftFormat->setBorderTop(borderStyle);
	topLeftFormat->setBorderLeft(borderStyle);
	topLeftFormat->setBorderTopColor(borderColor);
	topLeftFormat->setBorderLeftColor(borderColor);

	topRightFormat->setBorderTop(borderStyle);
	topRightFormat->setBorderRight(borderStyle);
	topRightFormat->setBorderTopColor(borderColor);
	topRightFormat->setBorderRightColor(borderColor);

	bottomLeftFormat->setBorderBottom(borderStyle);
	bottomLeftFormat->setBorderLeft(borderStyle);
	bottomLeftFormat->setBorderBottomColor(borderColor);
	bottomLeftFormat->setBorderLeftColor(borderColor);

	bottomRightFormat->setBorderBottom(borderStyle);
	bottomRightFormat->setBorderRight(borderStyle);
	bottomRightFormat->setBorderBottomColor(borderColor);
	bottomRightFormat->setBorderRightColor(borderColor);

	// 위쪽 테두리 (모서리 제외)
	for (int col = startCol + 1; col < endCol; ++col)
	{
		Format* cellFormat = sheet->cellFormat(startRow, col);
		if (cellFormat) topFormat->setFont(cellFormat->font());
		sheet->setCellFormat(startRow, col, topFormat);
	}

	// 아래쪽 테두리 (모서리 제외)
	for (int col = startCol + 1; col < endCol; ++col)
	{
		Format* cellFormat = sheet->cellFormat(endRow, col);
		if (cellFormat) bottomFormat->setFont(cellFormat->font());
		sheet->setCellFormat(endRow, col, bottomFormat);
	}

	// 왼쪽 테두리 (모서리 제외)
	for (int row = startRow + 1; row < endRow; ++row)
	{
		Format* cellFormat = sheet->cellFormat(row, startCol);
		if (cellFormat) leftFormat->setFont(cellFormat->font());
		sheet->setCellFormat(row, startCol, leftFormat);
	}

	// 오른쪽 테두리 (모서리 제외)
	for (int row = startRow + 1; row < endRow; ++row)
	{
		Format* cellFormat = sheet->cellFormat(row, endCol);
		if (cellFormat) rightFormat->setFont(cellFormat->font());
		sheet->setCellFormat(row, endCol, rightFormat);
	}

	// 모서리 셀 처리
	Format* cellFormat;

	// 왼쪽 위 모서리
	cellFormat = sheet->cellFormat(startRow, startCol);
	if (cellFormat) topLeftFormat->setFont(cellFormat->font());
	sheet->setCellFormat(startRow, startCol, topLeftFormat);

	// 오른쪽 위 모서리
	cellFormat = sheet->cellFormat(startRow, endCol);
	if (cellFormat) topRightFormat->setFont(cellFormat->font());
	sheet->setCellFormat(startRow, endCol, topRightFormat);

	// 왼쪽 아래 모서리
	cellFormat = sheet->cellFormat(endRow, startCol);
	if (cellFormat) bottomLeftFormat->setFont(cellFormat->font());
	sheet->setCellFormat(endRow, startCol, bottomLeftFormat);

	// 오른쪽 아래 모서리
	cellFormat = sheet->cellFormat(endRow, endCol);
	if (cellFormat) bottomRightFormat->setFont(cellFormat->font());
	sheet->setCellFormat(endRow, endCol, bottomRightFormat);

	// 내부 테두리 지우기
	for (int row = startRow + 1; row < endRow; ++row)
	{
		for (int col = startCol + 1; col < endCol; ++col)
		{
			cellFormat = sheet->cellFormat(row, col);
			if (cellFormat)
			{
				clearFormat->setFont(cellFormat->font());
				sheet->setCellFormat(row, col, clearFormat);
			}
		}
	}
}

void draw_all_borders(Book* book, Sheet* sheet, int startRow, int startCol, int endRow, int endCol, BorderStyle borderStyle, Color borderColor)
{
	if (!book || !sheet || startRow > endRow || startCol > endCol)
		return;  // 유효성 검사

	for (int row = startRow; row <= endRow; ++row)
	{
		for (int col = startCol; col <= endCol; ++col)
		{
			// 기존 셀의 포맷을 가져옵니다.
			Format* cellFormat = sheet->cellFormat(row, col);
			if (!cellFormat)
			{
				cellFormat = book->addFormat();
			}
			else
			{
				// 기존 포맷을 복사하여 새 포맷을 만듭니다.
				cellFormat = book->addFormat(cellFormat);
			}

			// 테두리 스타일을 설정합니다.
			cellFormat->setBorder(BORDERSTYLE_THIN);  // 모든 방향의 테두리를 설정
			cellFormat->setBorderColor(borderColor);

			// 수정된 포맷을 셀에 적용합니다.
			sheet->setCellFormat(row, col, cellFormat);
		}
	}
}
void write_global_env(Book* book, Sheet* sheet, GLOBAL_ENV* pGlobalEnv) {
	
	int posX = 0, posY = 0;// Adjust position for 0-based indexing

	//draw_outer_border(book, sheet, posY, 0, posY + iHeight - 1, iWidth, BORDERSTYLE_THIN, COLOR_BLACK);

	// First row settings
	sheet->writeStr(1, 0, L"GlobalEnv");
	
	posY = 2;
	sheet->writeStr(posY, posX++, L"SimulationWeeks");
	sheet->writeNum(posY, posX++, pGlobalEnv->SimulationWeeks);
	sheet->writeStr(posY, posX++, L"maxWeek");
	sheet->writeNum(posY, posX++, pGlobalEnv->maxWeek);
	sheet->writeStr(posY, posX++, L"WeeklyProb");
	sheet->writeNum(posY, posX++, pGlobalEnv->WeeklyProb);
	sheet->writeStr(posY, posX++, L"Hr_Init_H");
	sheet->writeNum(posY, posX++, pGlobalEnv->Hr_Init_H);
	sheet->writeStr(posY, posX++, L"Hr_Init_M");
	sheet->writeNum(posY, posX++, pGlobalEnv->Hr_Init_M);
	sheet->writeStr(posY, posX++, L"Hr_Init_L");
	sheet->writeNum(posY, posX++, pGlobalEnv->Hr_Init_L);
	sheet->writeStr(posY, posX++, L"Hr_LeadTime");
	sheet->writeNum(posY, posX++, pGlobalEnv->Hr_LeadTime);
	sheet->writeStr(posY, posX++, L"Cash_Init");
	sheet->writeNum(posY, posX++, pGlobalEnv->Cash_Init);
	sheet->writeStr(posY, posX++, L"ProblemCnt");
	sheet->writeNum(posY, posX++, pGlobalEnv->ProblemCnt);
	sheet->writeStr(posY, posX++, L"ExpenseRate");
	sheet->writeNum(posY, posX++, pGlobalEnv->ExpenseRate);
	sheet->writeStr(posY, posX++, L"selectOrder");
	sheet->writeNum(posY, posX++, pGlobalEnv->selectOrder);
	sheet->writeStr(posY, posX++, L"recruit");
	sheet->writeNum(posY, posX++, pGlobalEnv->recruit);
	sheet->writeStr(posY, posX++, L"layoff");
	sheet->writeNum(posY, posX++, pGlobalEnv->layoff);
	sheet->writeStr(posY, posX++, L"recruitTerm");
	sheet->writeNum(posY, posX++, pGlobalEnv->recruitTerm);

	// Define and apply styles
	Format* format = book->addFormat();
	format->setBorder(BorderStyle::BORDERSTYLE_THIN);
	format->setBorderColor(COLOR_BLACK);

	// Apply format to cells (assuming you want to apply to 1st row as an example)
	for (int col = 0; col < posX; ++col) {
		sheet->setCellFormat(2, col , format);
	}
}

void read_global_env(Book* book, Sheet* sheet, GLOBAL_ENV* pGlobalEnv) {
	if (!book || !sheet || !pGlobalEnv) return;

	int posX = 1;  // Start from the second column (because column 0 contains the field names)
	int posY = 2;  // Start from the row where the data starts (row 2)

	// Read the values and populate the structure
	pGlobalEnv->SimulationWeeks	= sheet->readNum(posY, posX );
	pGlobalEnv->maxWeek			= sheet->readNum(posY, posX += 2);
	pGlobalEnv->WeeklyProb		= sheet->readNum(posY, posX += 2);
	pGlobalEnv->Hr_Init_H		= sheet->readNum(posY, posX += 2);
	pGlobalEnv->Hr_Init_M		= sheet->readNum(posY, posX += 2);
	pGlobalEnv->Hr_Init_L		= sheet->readNum(posY, posX += 2);
	pGlobalEnv->Hr_LeadTime		= sheet->readNum(posY, posX += 2);
	pGlobalEnv->Cash_Init		= sheet->readNum(posY, posX += 2);
	pGlobalEnv->ProblemCnt		= sheet->readNum(posY, posX += 2);
	pGlobalEnv->ExpenseRate		= sheet->readNum(posY, posX += 2);
	pGlobalEnv->selectOrder		= sheet->readNum(posY, posX += 2);
	pGlobalEnv->recruit			= sheet->readNum(posY, posX += 2);
	pGlobalEnv->layoff			= sheet->readNum(posY, posX += 2);
	pGlobalEnv->recruitTerm		= sheet->readNum(posY, posX += 2);	
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
