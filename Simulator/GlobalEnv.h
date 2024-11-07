#pragma once

#include <vector>
#include <stdexcept> // For exception handling
#include <algorithm>

#include "libxl.h"
#include "hidden_libxl.h"
using namespace libxl;

#define MAX_CANDIDATES 50
#define MAX_DOING_TABLE_SIZE 10 //doing, done, deffer table 세로크기

// Order Tabe index for easy reference
enum OrderIndex {
	ORDER_SUM = 0,
	ORDER_ORD,
	ORDER_COUNT // Total number of OrderTable
};

// HR Tabe index for easy reference
enum HRIndex {
	HR_HIG = 0,
	HR_MID,
	HR_LOW,
	HR_COUNT // Total number of HR Table
};

int PoissonRandom(double lambda);

// project 에서 사용
#define MAX_N_CF  3
#define MAX_PRJ_TYPE 5
#define MAX_ACT  4

#define RND_HR_H  20
#define RND_HR_M  70


// activity의 타입에 대한 구조체
// activity_struct 시트의 cells(3,2) ~ cells(7,14)의 값으로 채워진다.
struct ACT_TYPE {
	int occurrenceRate;     // 타입별 발생 확률 (%)
	int cumulativeRate;     // 누적 확률 (%)
	int minPeriod;          // 최소 기간
	int maxPeriod;          // 최대 기간
	int patternCount;       // 패턴 수
	// 반복되는 패턴 번호와 확률
	int patterns[4][2];     // 최대 5개의 패턴 번호와 확률을 저장하는 2차원 배열
}; 
union ALL_ACT_TYPE {
	ACT_TYPE actTypes[5];  // 5개의 타일 발생 데이터를 위한 구조체 배열
	int asIntArray[MAX_PRJ_TYPE][sizeof(ACT_TYPE) / sizeof(int)];  // 5개의 타일 데이터를 정수 배열로 접근 (2차원 배열)
};


//////////////////////////////////////////////////////////////////////////
// activity_struct 시트의 cells(15,2) ~ cells(20,27)의 값으로 채워진다.
// 각 활동의 기간 비율과 인력 비율 패턴에 대한 구조체 정의
struct ACT_PATTERN {
	int minDurationRate;   // 최소 기간 비율 (%)
	int maxDurationRate;   // 최대 기간 비율 (%)
	int highHR;            // 고 인력 비율 (%)
	int mediumHR;          // 중 인력 비율 (%)
	int lowHR;             // 초 인력 비율 (%)
};

// 모든 활동의 패턴을 포함하는 구조체 정의
struct ALL_ACT_PATTERN {
	int patternCount;    // 활동 패턴 갯수
	ACT_PATTERN patterns[5];  // 5개의 활동 패턴
} ;

union ALL_ACTIVITY_PATTERN {
	ALL_ACT_PATTERN pattern[6];  // 6개의 활동 패턴을위한 구조체 배열
	int asIntArray[6][sizeof(ALL_ACT_PATTERN) / sizeof(int)];  // 6개의 활동 데이터를 정수 배열로 접근 (2차원 배열)
};


// Company 에서 사용
typedef struct _ACTIVITY {
	int activityType;  // 활동 유형
    int duration;      // 활동 기간
    int startDate;     // 시작 날짜
    int endDate;       // 종료 날짜
    int highSkill;     // 높은 기술 수준 인력 수
    int midSkill;      // 중간 기술 수준 인력 수
    int lowSkill;      // 낮은 기술 수준 인력 수
} ACTIVITY, *PACTIVITY;

// Sheet enumeration for easy reference
enum SheetName {
	WS_NUM_PARAMETERS = 0,
	WS_NUM_DASHBOARD,
	WS_NUM_PROJECT,
	WS_NUM_ACTIVITY_STRUCT,
	WS_NUM_DEBUG_INFO,
	WS_NUM_SHEET_COUNT // Total number of sheets
};

struct GLOBAL_ENV {
	int		SimulationWeeks;
	int		maxWeek; //최대 80주(18개월)간 진행되는 프로젝트를 시뮬레이션 마지막에 기록할 수도 있다.	
	double	WeeklyProb;
	int		Hr_Init_H;
	int		Hr_Init_M;
	int		Hr_Init_L;
	int		Hr_LeadTime;
	int		Cash_Init;
	int		ProblemCnt;	
	double	ExpenseRate;	// 비용계산에 사용되는 제경비 비율
	//double	profitRate;		// 프로젝트 총비용 계산에 사용되는 제경비 비율

	int		selectOrder;	// 선택 순서  1: 먼저 발생한 순서대로 2: 금액이 큰 순서대로 3: 금액이 작은 순서대로
	int		recruit;		// 충원에 필요한 운영비 (몇주분량인가?)
	int		layoff;			// 감원에 필요한 운영비 (몇주분량인가?)
	int		recruitTerm;	// 몇주마다 충원 감원 여부를 체크 할건가
};

////////////////////////////////////////////////////////////////////
// 프로젝트 속성
struct PROJECT {
	int category;		// 프로젝트 분류 (0: 외부 / 1: 내부)
	int ID;			// 프로젝트의 번호
	int orderDate;	// 발주일
	int startAvail;	// 시작 가능일
	int runningWeeks;		// 진행 개월수 (0: 미진행, 나머지: 진행한 기간)
	int experience;	// 경험 (0: 무경험 1: 유경험)
	int winProb;		// 성공 확률
	int nCashFlows;	// 비용 지급 횟수
		
	int endDate;		// 프로젝트 종료일
	int duration;		// 프로젝트의 총 기간

	int profit;	// 총 기대 수익 (HR 종속)

	// 현금 흐름
	int cashFlows[MAX_N_CF];	// 용역비를 받는 비율을 기록하는 배열
	int firstPay;		// 선금 액수
	int secondPay;		// 2차 지급 액수
	int finalPay;		// 3차 지급 액수
	int firstPayMonth;	// 선금 지급일
	int secondPayMonth;	// 2차 지급일
	int finalPayMonth;	// 3차 지급일

	// 활동
	int numActivities;          // 총 활동 수//    std::array<Activity, MAX_ACT> m_activities; // 활동에 관한 정보를 기록하는 배열
	ACTIVITY activities[MAX_ACT]; // 활동에 관한 정보를 기록하는 배열

	// 참고용변수
	int projectType;		// activity_struct 시트의 어느 타입의 프로젝트인가
	int activityPattern;	// activity_struct 시트의 어느 패턴인가
};



class Dynamic2DArray {
private:
	std::vector<std::vector<int>> data;

public:
	Dynamic2DArray() {}

	Dynamic2DArray(const Dynamic2DArray& other) : data(other.data) {} // 복사 생성자

	Dynamic2DArray& operator=(const Dynamic2DArray& other) { // 할당 연산자
		if (this != &other) {
			data = other.data;
		}
		return *this;
	}

	void Clear(int fill) {
		for (auto& row : data) {
			std::fill(row.begin(), row.end(), fill);
		}
	}

	class Proxy {
	private:
		Dynamic2DArray& array;
		int rowIndex;

	public:
		Proxy(Dynamic2DArray& arr, int index) : array(arr), rowIndex(index) {}

		int& operator[](int colIndex) {
			if (colIndex >= array.data[rowIndex].size()) {
				array.data[rowIndex].resize(colIndex + 1, -1);
			}
			return array.data[rowIndex][colIndex];
		}
	};

	Proxy operator[](int rowIndex) {
		if (rowIndex >= data.size()) {
			data.resize(rowIndex + 1);
		}
		return Proxy(*this, rowIndex);
	}

	void Resize(int x, int y) {
		data.resize(x);
		for (int i = 0; i < x; ++i) {
			data[i].resize(y, -1);
		}
	}

	void copyFromContinuousMemory(int* src, int rows, int cols) {
		Resize(rows, cols);
		for (int i = 0; i < rows; ++i) {
			for (int j = 0; j < cols; ++j) {
				data[i][j] = src[i * cols + j];
			}
		}
	}

	// Copy data to an external continuous memory block
	void copyToContinuousMemory(int* dest, int maxElements) {
		int index = 0;
		for (int i = 0; i < data.size() && index < maxElements; ++i) {
			for (int j = 0; j < data[i].size() && index < maxElements; ++j) {
				dest[index++] = data[i][j];
			}
		}
	}

	int getRows() const {
		return data.size();
	}

	int getCols() const {
		if (!data.empty()) return data[0].size();
		return 0;
	}
};



class DynamicProjectArray {
private:
	std::vector<std::vector<PROJECT>> data;

public:
	DynamicProjectArray() {}

	DynamicProjectArray(const DynamicProjectArray& other) : data(other.data) {} // 복사 생성자

	DynamicProjectArray& operator=(const DynamicProjectArray& other) { // 할당 연산자
		if (this != &other) {
			data = other.data;
		}
		return *this;
	}

	class Proxy {
	private:
		DynamicProjectArray& array;
		int rowIndex;

	public:
		Proxy(DynamicProjectArray& arr, int index) : array(arr), rowIndex(index) {}

		PROJECT& operator[](int colIndex) {
			if (colIndex >= array.data[rowIndex].size()) {
				array.data[rowIndex].resize(colIndex + 1); // 필요한 만큼 열 확장
			}
			return array.data[rowIndex][colIndex];
		}
	};

	Proxy operator[](int rowIndex) {
		if (rowIndex >= data.size()) {
			data.resize(rowIndex + 1); // 필요한 만큼 행 확장
		}
		return Proxy(*this, rowIndex);
	}

	void Resize(int x, int y) {
		data.resize(x);
		for (int i = 0; i < x; ++i) {
			data[i].resize(y);
		}
	}

	int getRows() const {
		return data.size();
	}

	int getCols() const {
		if (!data.empty()) return data[0].size();
		return 0;
	}
};




/*
1. 시그니처 SONG1 / 2.전체크기
3. 환경변수
4. 엑티비티타입
5. 엑티비티패턴
6. 현황판
프로젝트갯수
프로젝트
*/

#define SIGNITURE		{'A','H','N','1'} //pack 사용하지 않게 4바이트 정렬
#define TYPE_UNKNOWN		0
#define TYPE_ANH			1
#define TYPE_ENVIRONMENT	2
#define TYPE_ACTIVITY		3
#define TYPE_PATTERN		4
#define TYPE_ORDER			5
#define TYPE_DASHBD			6
#define TYPE_PROJECT		7

struct SAVE_SIG
{
	char signitre[4] = SIGNITURE;
	int totalLen;
};

struct SAVE_TL
{
	int type;
	int length;
};


bool OpenFile(const CString& filename, const TCHAR* mode, FILE** fp);
void CloseFile(FILE** fp);
ULONG WriteDataWithHeader(FILE* fp, int type, const void* data, size_t dataSize);
bool ReadDataWithHeader(FILE* fp, void* data, size_t expectedSize, int expectedType);


// 확률에 따라서 0 또는 1 생성
int ZeroOrOneByProb(int probability);

// 랜덤 숫자 생성 함수
int RandomBetween(int low, int high);

void clearSheet(Sheet* sheet);

void write_project_header(Book* book, Sheet* sheet);
void write_project_body(Book* book, Sheet* sheet, PROJECT* pProject);
void draw_outer_border(Book* book, Sheet* sheet, int startRow, int startCol, int endRow, int endCol, BorderStyle borderStyle, Color borderColor);
void draw_all_borders(Book* book, Sheet* sheet, int startRow, int startCol, int endRow, int endCol, BorderStyle borderStyle, Color borderColor);
void write_global_env(Book* book, Sheet* sheet, GLOBAL_ENV* pGlobalEnv);
void read_global_env(Book* book, Sheet* sheet, GLOBAL_ENV* pGlobalEnv);
void read_project_body(Book* book, Sheet* sheet, PROJECT* pProject, int projectIndex);