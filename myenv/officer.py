from person import Person


class Officer(Person):
    def __init__(self, rank, name, service_number, unit, position, user_type):
        # Officer 클래스는 추가 속성 없이 Person 클래스를 그대로 상속
        super().__init__(rank, name,  service_number, unit, position, user_type)

    def display_officer_info(self):
        """간부 고유의 정보를 출력하는 메서드"""
        return (f"계급: {self.rank}, 이름: {self.name},  군번: {self.service_number}, "
                f"소속: {self.unit}, 직책: {self.position}, 유저 타입: {self.user_type}")
