using System;
using System.Linq;

namespace Main
{
    public class CommandLineUtil
    {
        /// <summary>
        /// Get the specified command-line option's value.
        /// ex.
        ///   example.exe -d:C:\dir\
        ///   GetCommandOption("-d:") return "C:\dir\"
        /// </summary>
        /// <param name="opt">Prefix of the option (case sensitive)</param>
        /// <param name="def">String to return if the specified option does not exist.</param>
        /// <returns>It returns the result of removing the prefix from the options.</returns>
        public static string GetCommandOption(string opt, string def)
        {
            var result = def;
            var args = Environment.GetCommandLineArgs();
            if (args.Length > 1)
            {
                var o = args.Skip(1).Where(x => (x.IndexOf(opt) >= 0)).FirstOrDefault();
                if (o != null)
                {
                    result = o.Substring(opt.Length).Trim("\"".ToCharArray());
                }
            }
            return result;
        }

        /// <summary>
        /// Check existence of command line option
        /// </summary>
        /// <param name="opt">Option to confirm (Not case sensitive)</param>
        /// <returns>If exists option, then return true.</returns>
        public static bool HasCommandOption(string opt)
        {
            return Environment.GetCommandLineArgs().Skip(1).Any(x => x.ToLower() == opt.ToLower());
        }

    }
}
