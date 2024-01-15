using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using week08.Entities;

namespace week08.Abstractions
{
    public interface IAccountManager
    {
        BindingList<Account> Accounts { get; }

        Account CreateAccount(Account account);
    }
}
