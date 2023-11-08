package org.example;

//CopyBase.main(args) вызывает метод main из класса CopyBase для создания
//копии файла "Base table.xlsx" и сохранения его с именем "All.xlsx" в директории "Excel".
//CompareTables.main(args) вызывает метод main из класса CompareTables для сравнения таблиц в файле "All.xlsx".
//DeleteErrors.main(args) вызывает метод main из класса DeleteErrors для удаления некорректных данных в файле "All.xlsx".

public class Main {
    public static void main(String[] args) {
        CopyBase.main(args);
        CompareTables.main(args);
        DeleteErrors.main(args);
    }
}
