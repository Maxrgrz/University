import os
from pdf2docx import parse
from docx2pdf import convert
from PIL import Image

class OfficeTweaks:
    def __init__(self, path):
        self.path = path

    def show_menu(self):
        menu = [
            f" Текущий каталог: {self.path}\n\n",
            "Выберите действие\n\n"
            " 0. Сменить рабочий каталог\n",
            "1. Преобразовать PDF в Docx\n",
            "2. Преобразовать Docx в PDF\n",
            "3. Произвести сжатие изображений\n",
            "4. Удалить группу файлов\n",
            "5. Выход\n"
        ]
        print(*menu)

    def change_dir(self):
        while True:
            new_path = input("Укажите корректный путь к рабочему каталогу: ")
            if os.path.exists(new_path):
                self.path = new_path
                break
            else:
                print("Неверный адрес каталога, попробуйте снова")

        print(f"Текущий каталог: {self.path}\n")

    def pdf2docx(self):

        print("Список файлов с расширением .pfd в данном каталоге\n")
        files = list(filter(lambda x: x.endswith(".pdf"), os.listdir(self.path)))

        for n, i in enumerate(files, start=1):
            print(f"{n}. {i}")

        while True:
            choice = int(input("Введите номер файла для преобразования (чтобы преобразовать все файлы из данного каталога введите 0): "))
            if 0 < choice <= len(files):
                path = os.path.join(self.path, files[choice - 1])
                output_path = path[:-4] + ".docx"
                parse(path, output_path)  # Исправлено с convert на parse
                print(f"Файл {files[choice - 1]} преобразован в {output_path}")
                break
            elif choice == 0:
                for i in files:
                    path = os.path.join(self.path, i)
                    output_path = path[:-4] + ".docx"
                    parse(path, output_path)  # Исправлено с convert на parse
                    print("Файлы были успешно преобразованы")
                break
            else:
                print("Неверный ввод, попробуйте ещё раз")

    def docx2pdf(self):
        print("Список файлов с расширением .docx в данном каталоге\n")
        files = list(filter(lambda x: x.endswith(".docx"), os.listdir(self.path)))

        for n, i in enumerate(files, start=1):
            print(f"{n}. {i}")

        while True:
            choice = int(input("Введите номер файла для преобразования (чтобы преобразовать все файлы из данного каталога введите 0): "))
            if 0 < choice <= len(files):
                path = os.path.join(self.path, files[choice - 1])
                output_path = path[:-5] + ".pdf"
                convert(path, output_path)
                break
            elif choice == 0:
                for i in files:
                    path = os.path.join(self.path, i)
                    output_path = path[:-5] + ".pdf"
                    convert(path, output_path)
                break
            else:
                print("Неверный ввод, попробуйте ещё раз")

    def compress_image(self):
        print("Список файлов с расширением .jpeg, .gif, .png, .jpg в данном каталоге\n")
        files = list(filter(lambda x: x.endswith(".jpeg") or x.endswith(".gif") or x.endswith(".png") or x.endswith(".jpg"), os.listdir(self.path)))

        for n, i in enumerate(files, start=1):
            print(f"{n}. {i}")

        while True:
            choice = int(input("Введите номер файла для преобразования (чтобы преобразовать все файлы из данного каталога введите 0): "))
            if 0 < choice <= len(files):
                qual = int(input("Введите параметры сжатия (от 0% до 100%): "))
                path = os.path.join(self.path, files[choice - 1])
                image_file = Image.open(path)
                image_file.save(path, quality=qual)
                break
            elif choice == 0:
                qual = int(input("Введите параметры сжатия (от 0% до 100%): "))
                for i in files:
                    path = os.path.join(self.path, i)
                    image_file = Image.open(path)
                    image_file.save(path, quality=qual)
                    image_file.close()
                break
            else:
                print("Неверный ввод, попробуйте ещё раз")

    def del_group(self):
        menu = [
            " 1. Удалить все файлы, начинающиеся на определённую подстроку\n",
            "2. Удалить все файлы, заканчивающиеся на определённую подстроку\n",
            "3. Удалить все файлы, содержащие определённую подстроку\n",
            "4. Удалить все файлы по расширению\n",
        ]
        print(*menu)

        while True:
            choice = int(input("Выберите опцию: "))
            if choice == 1:
                substring = input("Введите подстроку: ")
                files = list(filter(lambda x: x.startswith(substring), os.listdir(self.path)))
                for i in files:
                    path = os.path.join(self.path, i)
                    os.remove(path)
                    print(f"Файл {i} успешно удалён!\n")
                break
            elif choice == 2:
                substring = input("Введите подстроку: ")
                files = list(filter(lambda x: x.endswith(substring), os.listdir(self.path)))
                for i in files:
                    path = os.path.join(self.path, i)
                    os.remove(path)
                    print(f"Файл {i} успешно удалён!\n")
                break
            elif choice == 3:
                substring = input("Введите подстроку: ")
                files = list(filter(lambda x: substring in x, os.listdir(self.path)))
                for i in files:
                    path = os.path.join(self.path, i)
                    os.remove(path)
                    print(f"Файл {i} успешно удалён!\n")
                break
            elif choice == 4:
                extension = input("Введите расширение: ")
                if not extension.startswith("."):
                    extension = "." + extension
                files = list(filter(lambda x: extension in x, os.listdir(self.path)))
                for i in files:
                    path = os.path.join(self.path, i)
                    os.remove(path)
                    print(f"Файл {i} успешно удалён!\n")
                break
            else:
                print("Неверный ввод, попробуйте ещё раз")

    def start(self):
        while True:
            self.show_menu()
            choice = input("Ваш выбор: ")
            if choice == "0":
                self.change_dir()
            elif choice == "1":
                self.pdf2docx()
            elif choice == "2":
                self.docx2pdf()
            elif choice == "3":
                self.compress_image()
            elif choice == "4":
                self.del_group()
            elif choice == "5":
                break

if __name__ == "__main__":
    Tweaks = OfficeTweaks(os.path.dirname(__file__))
    Tweaks.start()