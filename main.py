my_file = open("File.txt", "w+")
my_file.write("Файл!")
my_file.close()

print("Хотите ввести текст? Да/Нет")
a = input()
if a == "Да":
    my_file = open("File.txt", "a+")
    a = input()
    my_file.write(a)
    my_file.close()