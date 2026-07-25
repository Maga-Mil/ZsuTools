using System;
using NUnit.Framework;

namespace Tests
{
    [TestFixture]
    public class PersonTests
    {
        [Test]
        public void TestParsing()
        {
            var person1 = new ZsuTools.Entities.Person("КОВАЛЕНКО Іван Петрович");
            Assert.AreEqual("Іван", person1.FirstName);
            Assert.AreEqual("Петрович", person1.Patronymic);
            Assert.AreEqual("КОВАЛЕНКО", person1.LastName);

            var person2 = new ZsuTools.Entities.Person("Коваленко   іван    Петрович");
            Assert.AreEqual("Іван", person2.FirstName);
            Assert.AreEqual("Петрович", person2.Patronymic);
            Assert.AreEqual("КОВАЛЕНКО", person2.LastName);

            var person3 = new ZsuTools.Entities.Person("Іван КОВАЛЕНКО");
            Assert.AreEqual("Іван", person3.FirstName);
            Assert.AreEqual(string.Empty, person3.Patronymic);
            Assert.AreEqual("КОВАЛЕНКО", person3.LastName);

            var person4 = new ZsuTools.Entities.Person("КОВАЛЕНКО Іван");
            Assert.AreEqual("Іван", person4.FirstName);
            Assert.AreEqual(string.Empty, person4.Patronymic);
            Assert.AreEqual("КОВАЛЕНКО", person4.LastName);
            
            var person5 = new ZsuTools.Entities.Person("КОВАЛЕНКО I.П.");
            Assert.AreEqual("I", person5.FirstName);
            Assert.AreEqual("П", person5.Patronymic);
            Assert.AreEqual("КОВАЛЕНКО", person5.LastName);
        }
        
        [Test]
        public void TestEquality()
        {
            var person1 = new ZsuTools.Entities.Person("Іван КОВАЛЕНКО");
            var person2 = new ZsuTools.Entities.Person("Іван КОВАЛЕНКО");
            Assert.IsTrue(person1.Equals(person2));
            
            person1 = new ZsuTools.Entities.Person("КОВАЛЕНКО Іван");
            person2 = new ZsuTools.Entities.Person("Іван КОВАЛЕНКО");
            Assert.IsTrue(person1.Equals(person2));
            
            person1 = new ZsuTools.Entities.Person("КОВАЛЕНКО Іван Петрович");
            person2 = new ZsuTools.Entities.Person("Іван КОВАЛЕНКО");
            Assert.IsTrue(person1.Equals(person2));
        }

    }
}