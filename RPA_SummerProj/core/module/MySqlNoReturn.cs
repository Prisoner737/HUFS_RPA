using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Activities;
using RPA_SummerProj.core.implement;

namespace RPA_SummerProj.core.module
{

    public sealed class MySqlActivityEX : CodeActivity
    {
        // 형식 문자열의 작업 입력 인수를 정의합니다.
        public InArgument<string> serverName { get; set; }
        public InArgument<string> userName { get; set; }
        public InArgument<string> databaseName { get; set; }
        public InArgument<string> portNumber { get; set; }
        public InArgument<string> passWord { get; set; }
        public InArgument<string> sqlCommand { get; set; }

        // 작업 결과 값을 반환할 경우 CodeActivity<TResult>에서 파생되고
        // Execute 메서드에서 값을 반환합니다.
        protected override void Execute(CodeActivityContext context)
        {
            // 텍스트 입력 인수의 런타임 값을 가져옵니다.
            string server = context.GetValue(this.serverName);
            string user = context.GetValue(this.userName);
            string db = context.GetValue(this.databaseName);
            string port = context.GetValue(this.portNumber);
            string password = context.GetValue(this.passWord);
            string sql = context.GetValue(this.sqlCommand);

            MySqlManager myManager = new MySqlManager(server, user, db, port, password);
            myManager.MySqlEXCommand(sql);
        }
    }
}
