from person import Person


class Soldier(Person):
    def __init__(self, name, rank, service_number, unit, position, user_type, discharge_date):
        # Soldier만의 추가 속성인 discharge_date을 정의
        super().__init__(name, rank, service_number, unit, position, user_type)
        self.discharge_date = discharge_date  # 병사의 추가 속성: 근무 형태

    def display_soldier_info(self):
        """병사 고유의 정보를 출력하는 메서드"""
        return (f"계급: {self.rank}, 이름: {self.name},  군번: {self.service_number}, "
                f"소속: {self.unit}, 직책: {self.position}, 유저 타입: {self.user_type}, 전역일: {self.discharge_date}")
