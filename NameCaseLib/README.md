Библиотека NameCaseLib склонения ФИО для .NET
==================

Поддерживаются русский и украинский языки.

Пример использования:
```csharp
  var surname = "Иванов";
  var name = "Иван";
  var patronymic = "Иванович";
  var fullName = $"{surname} {name} {patronymic}";
  // Создается экземпляр класса для дальнейшей работы
  var nameCaseLibInstance = new Ru();
  // Определяется пол по ФИО
  var gender = nameCaseLibInstance.DetectGender(fullName);
  // Склонение фамилии
  var surnames = nameCaseLibInstance.QSurname(surname, gender);
  // Склонение имени
  var names = nameCaseLibInstance.QName(name, gender);
  // Склонение отчества
  var patronymics = nameCaseLibInstance.QFatherName(patronymic, gender);
  // Склонение ФИО
  var fullNames = nameCaseLibInstance.QFullName(fullName, gender);
```
