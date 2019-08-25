using System.Activities;
using RPA_SummerProj.core.implement;

namespace RPA_SummerProj.core.module
{

    public sealed class OutlookShowUnReadMSG : CodeActivity
    {
        // 작업 결과 값을 반환할 경우 CodeActivity<TResult>에서 파생되고
        // Execute 메서드에서 값을 반환합니다.
        protected override void Execute(CodeActivityContext context)
        {
            MailManager myMail = new MailManager();
            myMail.receiveMail();
        }
    }
}
