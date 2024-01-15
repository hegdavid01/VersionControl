using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitTestExample.Test
{
    public class AccountControllerTestFixture
    {
        [
       Test,
       TestCase("abcd1234", false),
       TestCase("irf@uni-corvinus", false),
       TestCase("irf.uni-corvinus.hu", false),
       TestCase("irf@uni-corvinus.hu", true)
       ]

        public void TestValidateEmail(string email, bool expectedResult)
        {
            var accountController = new AccountController();
            var actualResult = accountController.ValidateEmail(email);
            Assert.AreEqual(expectedResult, actualResult);
        }
           [Test,
           TestCase("asdfASDF", false),
           TestCase("ASDF1234", false),
           TestCase("asdf1234", false),
           TestCase("asd123", false),
           TestCase("Asdf1234", true),
           ]

    }   
}
