using System.ComponentModel.DataAnnotations;

namespace DictionaryManagement_Models.IntDBModels
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class CheckControlSymbols : ValidationAttribute
    {
        public override bool IsValid(object? value)
        {
            var isValid = true;
            if (value != null)
            {
                var inputValue = value as string;
                if (inputValue != null)
                {
                    foreach (char c in inputValue)
                    {
                        if (Char.IsControl(c))
                        {
                            isValid = false;
                            break;
                        }
                    }
                }
            }
            if (isValid != true)
                this.ErrorMessage = "В поле присутствуют непечатные символы";
            return isValid;
        }
    }
}
