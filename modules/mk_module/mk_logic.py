from .mk_calculator import calculate_time

def prepare_operations_data(ops_list):
    prepared_list = []
    for op in ops_list:
        # Если времени нет в базе, считаем его по формуле (пока заглушка)
        default_time = 5.0

        op_data = {
            "name": op["name"],
            "equipment": op["equipment"],
            "code": op.get("code", ""),
            "description": op.get("description", ""),
            "tool": op.get("tool", ""),
            "profession": op.get("profession", "Слесарь"),
            "rank": op.get("rank", "3"),
            "time_Tsht": op.get("time_Tsht", str(default_time))
        }
        prepared_list.append(op_data)
    return prepared_list