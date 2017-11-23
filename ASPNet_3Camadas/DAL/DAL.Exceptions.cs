using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;

namespace DAL.Exceptions
{
    /// <summary>
    /// Classe personalizda para erros de conexão
    /// </summary>
    [Serializable]
    internal class DALExceptionConnectionError : Exception
    {
        public DALExceptionConnectionError(){}
        public DALExceptionConnectionError(string message) : base(message){}
        public DALExceptionConnectionError(string message, Exception innerException) : base(message, innerException) {}
        protected DALExceptionConnectionError(SerializationInfo info, StreamingContext context) : base(info, context){}
    }
    /// <summary>
    /// Exceção especifica para erros na abertura da conexão
    /// </summary>
    [Serializable]
    internal class DALExceptionConnectionOpen : Exception
    {
        public DALExceptionConnectionOpen(){}
        public DALExceptionConnectionOpen(string message) : base(message){}
        public DALExceptionConnectionOpen(string message, Exception innerException) : base(message, innerException){}
        protected DALExceptionConnectionOpen(SerializationInfo info, StreamingContext context) : base(info, context){}
    }


    [Serializable]
    internal class DALExceptionExecuteReader : Exception
    {
        public DALExceptionExecuteReader()
        {
        }

        public DALExceptionExecuteReader(string message) : base(message)
        {
        }

        public DALExceptionExecuteReader(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected DALExceptionExecuteReader(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }


    [Serializable]
    internal class DALExceptionCommand : Exception
    {
        public DALExceptionCommand()
        {
        }

        public DALExceptionCommand(string message) : base(message)
        {
        }

        public DALExceptionCommand(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected DALExceptionCommand(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }

}
