import sqlite3


class BotDB:

    def __init__(self, db_file):
        self.conn = sqlite3.connect(db_file)
        self.cursor = self.conn.cursor()

    def user_exists(self, user_id):
        """Проверяем, есть ли заказчик в базе"""
        result = self.cursor.execute("SELECT `id` FROM `users` WHERE `user_id` = (?)", (user_id,))
        return not bool(result.fetchone() is None)

    def get_balance(self, user_id):
        """Достаем баланс в базе по его user_id"""
        result = self.cursor.execute("select balance from users where user_id = (?)", (user_id,))
        return result.fetchone()[0]

    def add_user(self, user_id: int, username):
        """Добавляем юзера в базу"""
        self.cursor.execute(
            "INSERT INTO users (user_id, username ) VALUES(?,?);",
            (int(user_id), username))
        return self.conn.commit()
