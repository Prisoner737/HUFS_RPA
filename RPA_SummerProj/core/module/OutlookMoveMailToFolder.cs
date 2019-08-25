﻿using System.Activities;
using RPA_SummerProj.core.implement;

namespace RPA_SummerProj.core.module
{

    public sealed class OutlookMoveMailToFolder : CodeActivity
    {
        // 형식 문자열의 작업 입력 인수를 정의합니다.
        public InArgument<string> folderName { get; set; }
        public InArgument<string> mailName { get; set; }

        // 작업 결과 값을 반환할 경우 CodeActivity<TResult>에서 파생되고
        // Execute 메서드에서 값을 반환합니다.
        protected override void Execute(CodeActivityContext context)
        {
            // 텍스트 입력 인수의 런타임 값을 가져옵니다.
            string tgtFolder = context.GetValue(this.folderName);
            string tgtMail = context.GetValue(this.mailName);
            MailManager myMail = new MailManager();
            myMail.moveMailFolder(tgtMail, tgtFolder);
        }
    }
}
