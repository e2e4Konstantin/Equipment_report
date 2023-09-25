


def glue_letter(letter: str, src_list: str) -> list[str]:
    """ К символу letter приклеивает каждый символ src_list  """
    return [f"{letter}{v}" for v in src_list]


def generate_column_names(required_size: int) -> list[str] | None:
    """ Создает список названий столбцов из букв английского алфавита. """
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    if 0 < required_size < 16384:
        if required_size <= len(alphabet):
            return [*alphabet][:required_size]
        else:
            entire = required_size // len(alphabet)
            extra = [v for x in alphabet[:entire] for v in glue_letter(x, alphabet)]
            return ([*alphabet] + extra)[:required_size]
    return None

if __name__ == "__main__":
    print(generate_column_names(26))
