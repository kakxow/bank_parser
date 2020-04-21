from typing import List

import win32com.client as win32


def _wrap_html(txt: str) -> str:
    html = (
        '<html>'
        '<head>'
        '<style>p {font-size: 11pt; font-family: "Calibri";}</style>'
        '</head>'
        f'<body>{txt}</body>'
        '</html>'
    )
    return html


class Mail:
    recepients: List[str]
    subject: str
    body: str
    copy_recepients: List[str]
    blind_copy_recepients: List[str]
    attachments_paths: List[str]

    def __init__(
        self,
        recepients: List[str],
        subject: str,
        body: str,
        copy_recepients: List[str] = None,
        blind_copy_recepients: List[str] = None,
        attachments_paths: List[str] = None
    ):
        self.recepients = recepients
        self.subject = subject
        self.body = body
        self.copy_recepients = copy_recepients or []
        self.blind_copy_recepients = blind_copy_recepients or []
        self.attachments_paths = attachments_paths or []

    def send_mail(self, importance: int = 1, send: bool = False) -> None:
        """
        importance = {'High': 2, 'Normal': 1, 'Low': 0,}
        """
        outlook = win32.Dispatch('outlook.application')

        mail = outlook.Application.CreateItem(0)
        body_without_sig = _wrap_html(self.body)
        mail.Display()
        signature = mail.HTMLBody
        for recepient in self.recepients:
            mail.Recipients.Add(recepient)
        mail.CC = ';'.join(self.copy_recepients)
        mail.BCC = ';'.join(self.blind_copy_recepients)
        mail.Subject = self.subject
        mail.HTMLBody = body_without_sig + '<br>' + signature
        for path in self.attachments_paths:
            mail.Attachments.Add(path)
        mail.importance = importance
        if send:
            mail.Send()
