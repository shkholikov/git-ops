interface IMailData {
    To: string | string[];
    CC: string | string[];
    Subject: string;
    Body: string
}

interface ISendMail {
//class shape
}

class SendMail implements ISendMail {
    private mailData: IMailData;
    private mailTemplates: string[];
    public isTestEnv: boolean;
    private flowEndPoint: string;

    constructor(mailData: IMailData, mailTemplates: string[]) {
        this.mailData = mailData
        this.mailTemplates = mailTemplates
        this.isTestEnv = true
    }

    private prepareMailBody(): void {
        //prepare mail body
        console.log(this.mailData.Body)
    }

    sendMail() {
        this.prepareMailBody()

        //send mail
        console.log(this.mailData)
        console.log(this.mailTemplates)


    }
}

const oops = new SendMail({
    To: ["imshakhzod@gmail.com"],
    CC: ["imshakhzod@gmail.com"],
    Subject: "MailBox",
    Body: "This is email body from you!"
}, ["EmailTemplate"])

oops.sendMail()