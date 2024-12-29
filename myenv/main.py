# import openpyxl
# from openpyxl import Workbook
import time
from datetime import datetime

from prettytable import PrettyTable

from officer import Officer
from person import Person
from soldier import Soldier


class DutyScheduler:
    def __init__(self):
        self.users = {}
        self.duty_type = []  # 근무 종류
        self.duty_order = {}  # 근무 순번표
        self.duty_date = ""
        

    def display_main_menu(self):
        menu_title = "메인 메뉴"
        menu_items = [
            "회원가입",
            "로그인",
            "유저 정보 조회",
            "유저 정보 수정",
            "근무표를 엑셀파일로 변환하기",
            "프로그램 종료"
        ]

        print("=" * 80)
        print()
        print(f"{menu_title:^80}")  # 제목 가운데 정렬
        print()
        print("=" * 80)
        print()

        for i, item in enumerate(menu_items, start=1):
            print(f" [{i}] {item}")

        print()
        print("=" * 80)

    def register_user(self, rank, name, service_number, unit, position, discharge_date, is_admin=False):
        if service_number in self.users:
            print("이미 등록된 군번입니다.")
            return

        if is_admin:
            user_type = "Admin"
            user = Officer(rank, name, service_number,
                           unit, position, user_type)
        else:
            user_type = "Normal User"
            user = Soldier(rank, name, service_number, unit,
                           position, user_type, discharge_date)

        self.users[service_number] = user
        print(f"{name} 등록 완료!")

    def show_users(self, users):
        user_table = PrettyTable()
        user_table.field_names = ["군번", "계정 권한", "소속", "직책명", "계급", "이름"]

        for _, user_info in users.items():
            user_table.add_row([user_info.service_number,
                                user_info.user_type,
                                user_info.unit,
                                user_info.position,
                                user_info.rank,
                                user_info.name])
        return user_table

    def show_duty_inputs(self, duties_order):
        duty_table = PrettyTable()
        duty_table.field_names = ["근무", "순번"]

        for duties_type, order in duties_order.items():
            duty_table.add_row([duties_type, order])

        return duty_table

    def login(self, service_number):
        user = self.users.get(service_number)
        if not user:
            print("사용자를 찾을 수 없습니다.")
            return None
        return user

    def display_admin_menu(self):
        menu_title = "관리자 메뉴"
        menu_items = [
            "근무표 날짜 입력하기",
            "근무 종류 입력하기",
            "근무 순번표 입력하기",
            "현재까지 입력된 값들 확인하기",
            "근무종류 초기화",
            "순번 초기화",
            "근무표 제작하기",
            "메뉴 나가기"
        ]

        print()
        print("=" * 80)
        print()
        print(f"{menu_title:^80}")  # 제목 가운데 정렬
        print("=" * 80)
        print()

        for i, item in enumerate(menu_items, start=1):
            print(f" [{i}] {item}")

        print()
        print("=" * 80)

    def assign_duties(self, start_date, end_date):
        current_order_index = 0
        current_date = start_date

        while current_date <= end_date:
            self.duty_schedule[current_date] = {}

            for duty in self.duties:
                personnel_count = 1  # Assume 1 person per duty; can be modified as needed
                assigned_personnel = []

                for _ in range(personnel_count):
                    assigned_personnel.append(
                        self.duty_order[current_order_index % len(self.duty_order)])
                    current_order_index += 1

                self.duty_schedule[current_date][duty] = assigned_personnel

            # Increment day (assumes integer format for simplicity)
            current_date += 1
    
    def create_excel_file(file_name):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"{self.duty_date.year}년 {self.duty_date.month}월 근무표" 
        self.wb.save(file_name)


def main_menu():
    scheduler = DutyScheduler()

    print("=" * 80)
    print()
    print("\n  🌟 반갑습니다! 근무 자동화 프로그램 Schedulizer입니다 🌟")
    print("  이 프로그램은 근무 배정의 효율성을 위해 설계되었습니다.")
    print("  아래 메뉴를 통해 회원 관리, 근무 설정 및 파일 변환 등을 진행하실 수 있습니다.")
    print("\n  ⚡ 시작하시려면 원하는 메뉴를 선택해주세요! ⚡\n")
    print()
    print("=" * 80)

    while True:
        print()
        scheduler.display_main_menu()
        print()
        choice = input("선택: ")
        print()

        # 회원가입
        if choice == "1":
            rank = input("계급: ")
            name = input("이름: ")
            service_number = input("군번 (예: 24-12345678): ")
            unit = input("소속 (예: OO대대 OO중대): ")
            position = input("직책: ")

            while True:
                try:
                    user_input = input("전역 예정일 (예: YYYY-MM-DD): ")
                    discharge_date = datetime.strptime(
                        user_input, "%Y-%m-%d").date()
                    break
                except ValueError:
                    print("잘못된 형식입니다. 'YYYY-MM-DD' 형식으로 입력해주세요.")

            while True:
                is_admin = input("관리자 계정입니까? (y/n): ").lower()
                if is_admin in ["y", "n"]:
                    is_admin = is_admin == "y"  # y는 True, n은 False로 변환
                    break
                else:
                    print("잘못된 입력입니다. 'y' 또는 'n'을 입력해주세요.")

            if is_admin == 'y':
                is_admin = True

            scheduler.register_user(
                rank, name, service_number, unit, position, discharge_date, is_admin)

        # 로그인
        elif choice == "2":
            service_number = input("로그인을 하려면 군번을 입력하세요: ")
            user = scheduler.login(service_number)
            print()
            print(f"{user.rank} {user.name}님 환영합니다.")

            if user:
                while True:
                    scheduler.display_admin_menu()
                    print()
                    admin_choice = int(input("선택: "))

                    if admin_choice == 1:
                        print("* 만약 기존에 입력 되어있던 날짜가 있다면 지금 입력한 날짜로 갱신됩니다.")
                        while True:
                            user_input = input("몇년 몇월 근무표인지 입력해주세요 (예: 2025-03): ")
                            print()
                            try:
                                scheduler.duty_date = datetime.strptime(
                                    user_input, "%Y-%m")
                                break
                            except ValueError:
                                print("잘못된 형식입니다. YYYY-MM 형식으로 다시 입력해주세요.")

                    elif admin_choice == 2:
                        print("=" * 80)
                        print()
                        print("불침번\nB1\n식기세척\n책임분대장\nCCTV\n여단 쓰레기 분리수거")
                        print()
                        print("=" * 80)

                        continue_input = True

                        while continue_input:
                            print()
                            print("엑셀에 입력할 순서대로 입력해주세요.")
                            user_input = input("입력: ")
                            print()
                            scheduler.duty_type.append(user_input)
                            while True:
                                continue_input = input(
                                "근무를 더 추가 하시겠습니까? (y/n): ").lower()
                                if continue_input in ["y", "n"]:
                                    continue_input = continue_input == "y"  # y는 True, n은 False로 변환
                                    break
                                else:
                                    print("잘못된 입력입니다. 'y' 또는 'n'을 입력해주세요.")

                    elif admin_choice == 3:
                        print("=" * 80)
                        print()
                        print("근무 순번표를 입력해주세요.")
                        print()
                        print("* 준비된 순번표 그대로 병사 이름을 차례대로 입력해주세요.")
                        print("* 이름을 하이픈(-)으로 구분해주세요.")
                        print("* 입력 예시: 홍길동-김민준-박현준-이민재-유동민")
                        print()
                        print("=" * 80)
                        print()

                        continue_input = True

                        while continue_input:
                            duty = input("어떤 근무의 순번을 입력하시겠습니까? (예: 불침번): ")
                            user_input = input("순번 입력: ")
                            scheduler.duty_order[duty] = user_input
                            continue_input = input(
                                "다른 근무의 순번을 추가로 입력하시겠습니까? (y/n): ").lower()
                            if continue_input in ["y", "n"]:
                                continue_input = continue_input == "y"  # y는 True, n은 False로 변환
                            else:
                                print("잘못된 입력입니다. 'y' 또는 'n'을 입력해주세요.")

                    elif admin_choice == 4:
                        print()
                        print("{:=^71}".format("현재까지 입력된 값들"))
                        print()

                        if len(str(scheduler.duty_date)) > 0:
                            print("[날짜] {}년 {}월 근무표 입니다".format(
                                scheduler.duty_date.year, scheduler.duty_date.month))
                        else:
                            print("[날짜] 등록된 날짜가 없습니다.")

                        print()
                        print("[순번] 현재 등록 되어있는 근무와 그에 따른 순번입니다.")
                        print(scheduler.show_duty_inputs(scheduler.duty_order))
                        print()
                        print("=" * 80)

                    elif admin_choice == 5:
                        scheduler.duty_type = []
                        print("입력되어있었던 근무 종류를 모두 초기화 완료하였습니다.")

                    elif admin_choice == 6:
                        scheduler.duty_order = {}
                        print("입력되어있었던 순번을 모두 초기화 완료하였습니다.")

                    elif admin_choice == 7:
                        print("근무표를 자동으로 제작합니다.")
                        print("Loading...")
                        time.sleep(5)

                    elif admin_choice == 8:
                        print("관리자 메뉴를 종료합니다.")
                        break

                    else:
                        print("잘못된 입력입니다. 다시 선택해주세요.")

        # 유저 정보 조회.
        elif choice == "3":
            print("{:=^69}".format("현재 등록되어있는 유저들"))
            print(scheduler.show_users(scheduler.users))
            print('=' * 80)

        # 유저 정보 수정.
        elif choice == "4":
            service_number = input("수정할 유저의 군번을 입력하세요: ")
            print("수정 가능한 항목: rank, name, unit, position")
            field = input("수정할 항목: ")
            value = input(f"새로운 {field}: ")
            scheduler.update_user(service_number, **{field: value})

        # 근무표를 엑셀파일로 변환하기.
        elif choice == "5":
            filename = input("저장할 파일 이름(예: duty_schedule.xlsx): ")
            scheduler.export_to_excel(filename)

        # 프로그램 종료
        elif choice == "6":
            print("프로그램을 종료합니다.")
            break

        else:
            print("잘못된 입력입니다. 다시 선택해주세요.")


if __name__ == "__main__":
    print("")
    main_menu()
