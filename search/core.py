import os
import re
import zipfile

from collections.abc import Generator

from compoundfiles import CompoundFileReader


def search_docx(filepath: str, string: str) -> bool:
    zdoc = zipfile.ZipFile(filepath, mode='r')
    xml_bytes = zdoc.read('word/document.xml')
    return string.encode() in xml_bytes


def search_doc(filepath: str, string: str) -> bool:
    with CompoundFileReader(filepath) as doc:
        if doc.root is None:
            raise Exception('A document root of "{}" is None'.format(filepath))
    
        is_found = False
        with doc.open(doc.root['WordDocument']) as f:

            f.seek(0x01A2)
            fc_clx = int.from_bytes(f.read(4), byteorder='little')
            lcb_clx = int.from_bytes(f.read(4), byteorder='little')

            f.seek(0x000A)
            table_flag = int.from_bytes(f.read(2), byteorder='little')
            table_num = (table_flag & 0x0200) == 0x0200
            table_name = '{0:d}Table'.format(table_num)

            with doc.open(doc.root[table_name]) as ft:
                ft.seek(fc_clx)
                clx = ft.read(lcb_clx)
            
            pos = 0
            while True:
                typeEntry = clx[pos]
                if typeEntry == 2:
                    lcb_piece_table = int.from_bytes(clx[pos + 1:pos + 2], byteorder='little')
                    piece_table = clx[pos + 5:]
                    if lcb_piece_table != len(piece_table):
                        pos = pos + 1 + 1 + clx[pos + 1]
                        continue
                    break
                elif typeEntry == 1:
                    pos = pos + 1 + 1 + clx[pos + 1]
                else:
                    break
            
            piece_count = (lcb_piece_table - 4) // 12

            for i in range(piece_count):
                s_idx = i * 4
                e_idx = (i + 1) * 4
                cp_start = int.from_bytes(piece_table[s_idx:s_idx + 4], byteorder='little')
                cp_end = int.from_bytes(piece_table[e_idx:e_idx + 4], byteorder='little')

                offset_pd = ((piece_count + 1) * 4) + (i * 8)
                piece_descriptor = piece_table[offset_pd:offset_pd + 8]

                fc_value = int.from_bytes(piece_descriptor[2:6], byteorder='little')
                is_ansi = (fc_value & 0x40000000) == 0x40000000
                fc = fc_value & 0xBFFFFFFF

                cb = cp_end - cp_start
                if not is_ansi:
                    encoding = 'utf-16-le'
                    cb *= 2
                else:
                    encoding = 'utf-8'
                    fc //= 2

                f.seek(cb)
                content_bytes = f.read(fc)
                if string.encode(encoding) in content_bytes:
                    is_found = True
                    break
    
    return is_found


def iter_docs(rootdir: str, string: str) -> Generator[tuple[str, bool], None, None]:
    try:
        for base, dirs, files in os.walk(rootdir):
            for file in files:
                __, extention = os.path.splitext(file)
                filepath = os.path.join(base, file)
                try:
                    if extention == '.docx':
                        is_present = search_docx(filepath, string)
                    elif extention == '.doc':
                        is_present = search_doc(filepath, string)
                    else:
                        continue
                except Exception as e:
                    print('[{filepath}]: ошибка {err}'.format(filepath=filepath, err=e))
                else:
                    yield filepath, is_present
    except GeneratorExit:
        pass
