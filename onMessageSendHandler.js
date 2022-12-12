/* global Office, fetch */

let _mailboxItem;
let mensagemVinculada;

function onMessageSendHandler(event) {
  _mailboxItem = Office.context.mailbox.item;

  _mailboxItem.subject.getAsync({ asyncContext: event }, function (asyncSubject) {
    throwError(asyncSubject, "Falha na recuperação do assunto: ");

    _mailboxItem.loadCustomPropertiesAsync(function (asyncCustomProp) {
      throwError(asyncCustomProp, "Falha na recuperação de propriedades customizadas: ");

      const tagRe = /\[[a-zA-Z ]+:.\d+\]/g;
      mensagemVinculada = tagRe.exec(asyncSubject.value);

      if (mensagemVinculada === null) {
        mensagemVinculada = asyncCustomProp.value.get("mensagemVinculada");
        if (
          !["", null, undefined].includes(mensagemVinculada) &&
          typeof mensagemVinculada === "string" &&
          tagRe.exec(mensagemVinculada) !== null
        ) {
          alteraAssunto(`${mensagemVinculada.toLocaleUpperCase()} ${asyncSubject.value}`);
        } else {
          asyncSubject.asyncContext.completed({ allowEvent: true });
          return;
        }
      } else {
        mensagemVinculada = mensagemVinculada[0].toLocaleLowerCase();
      }

      salvarMensagemEAnexos(asyncSubject);
    });
  });
}

const alteraAssunto = (novoAssunto) => {
  _mailboxItem.subject.setAsync(novoAssunto, function (asyncResult) {
    throwError(asyncResult, "Falha na alteração do assunto: ");
  });
};

const salvarMensagemEAnexos = (event) => {
  _mailboxItem.body.getAsync(Office.CoercionType.Html, (asyncBody) => {
    requisicaoSalvarMensagemEAnexos(event, { body: asyncBody.value });
  });
};

const requisicaoSalvarMensagemEAnexos = async (event, serviceRequest) => {
  const settings = Office.context.roamingSettings;

  const urlBase = settings.get("urlBase");
  const token = settings.get("token");

  const entidade = /[a-zA-Z]{3,}/g.exec(mensagemVinculada)[0];
  const id = /\d+/g.exec(mensagemVinculada)[0];

  fetch(`${urlBase}api/extensao-office/vincular-mensagem/${entidade}/${id}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization-Vios-Office": `Bearer ${token}`,
    },
    body: JSON.stringify(serviceRequest),
  })
    .then(async (response) => {
      const data = await response.json();
      if (!response.ok) {
        const error = (data && data.message) || response.status;
        return Promise.reject(error);
      }

      return data;
    })
    .then(() => {
      event.asyncContext.completed({ allowEvent: true });
    })
    .catch((error) => {
      event.asyncContext.completed({ allowEvent: false, errorMessage: "Error: " + error });
    });
};

const throwError = (asyncResult, mensagem = "") => {
  if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
    throw mensagem + JSON.stringify(asyncResult.error);
  }
};

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
