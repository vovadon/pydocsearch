import time

from core import iter_docs

searchtext = "гур"
rootdir = "C:\\test\\"


def main() -> None:
    print('Выполняется поиск файлов в директории {root}'.format(root=rootdir))

    start_time = time.time()
    total = 0
    for filepath, str_present in iter_docs(rootdir, searchtext):
        if str_present:
            print(filepath)
        total += 1
        # A stop condition
        # if total > 4:
        #     break # Will raise GeneratorExit

    elapsed = str(time.time() - start_time)

    print('Поиск завершен. Файлов проверено: {total}. Время выполнения: {time}'.format(total=total, time=elapsed))


if __name__ == '__main__':
    main()