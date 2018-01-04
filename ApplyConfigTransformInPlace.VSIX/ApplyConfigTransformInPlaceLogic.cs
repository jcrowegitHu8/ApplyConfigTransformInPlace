using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ApplyConfigTransformInPlace.VSIX
{
    public static class ApplyConfigTransformInPlaceLogic
    {
        private static Regex supportedTransformRegex = new Regex(@"([\w]+\.)[\w- ]+\.config", RegexOptions.IgnoreCase);
        public static bool IsSupportedTransform(string fileName, out string destinationPrefix)
        {
            destinationPrefix = string.Empty;
            var match = supportedTransformRegex.Match(fileName);
            if (match.Success)
            {
                Debug.WriteLine($"{fileName} is a supported transform");
                destinationPrefix = match.Groups[1].Value;
                return true;
            }

            Debug.WriteLine($"{fileName} is NOT a supported transform");
            return false;
        }


    }
}
