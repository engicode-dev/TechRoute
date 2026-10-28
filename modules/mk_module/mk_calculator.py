# modules/mk_module/mk_calculator.py

def calculate_time(operation_name, equipment, base_time=5.0):
    """
    Функция для расчета норм времени.
    В будущем здесь можно сделать учет режимов резания, материала и т.д.
    """
    # Пример: базовое время умножается на коэффициент в зависимости от станка
    coefficients = {
        "Станок 16К20": 1.2,
        "Станок 2Н135": 1.5,
        "Сварочный пост": 2.0,
        "Гильотина": 0.8
    }

    coef = coefficients.get(equipment, 1.0)
    t_sht = base_time * coef  # Штучное время

    # Возвращаем словарь с нормами
    return {
        "То": round(base_time * 0.6, 2),  # Основное время (60%)
        "Тв": round(base_time * 0.4, 2),  # Вспомогательное время (40%)
        "Тшт": round(t_sht, 2)  # Штучное время
    }