class Person:
    def __init__(self, rank, name, service_number, unit, position, user_type):
        self.name = name  # 이름
        self.rank = rank  # 계급
        self.service_number = service_number  # 군번
        self.unit = unit  # 소속
        self.position = position  # 직책
        self.user_type = user_type  # 유저 타입

    def display_info(self):
        """공통된 정보를 출력하는 메서드"""
        return (f"계급: {self.rank}, 이름: {self.name},  군번: {self.serial_number}, "
                f"소속: {self.unit}, 직책: {self.position}, 유저 타입: {self.user_type}")
