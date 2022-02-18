using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Automation.Ui.Repository.Enums
{
    /// <summary>
    /// Login Objects
    /// </summary>
    public enum LoginObjects
    {
        [Description("UserNameTextBox")]
        UserName,
        [Description("PasswordTextBox")]
        Password,
        [Description("LoginButton")]
        SubmitBtn
    }
}