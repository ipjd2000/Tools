using System;

namespace LLDOCX.DeepCheck
{
    /// <summary>A Deep Check failure whose Message is safe to show to the user.</summary>
    public class DeepCheckException : Exception
    {
        public DeepCheckException(string message, Exception inner = null) : base(message, inner) { }
    }
}
