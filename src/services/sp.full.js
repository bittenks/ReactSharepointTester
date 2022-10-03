const isProduction = process.env.NODE_ENV !== "development"
const url = isProduction ? "https://bayergroup.sharepoint.com/sites/022971/" : "http://localhost:8080/"

function CrudSharepoint(baseUrl) {
    var digestToken = '' // Variavel para armazenar o token do digest
    var globalScope = this // Utilizado em algumas funções que serão recursivas

    /**
     * @description Este event listener inicia a função getDigest ao finalizar o carregamento da página e executa a cada 60 segundos 
     * Para maiores informações leia https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/basics/working-with-requestdigest
     */
    document.addEventListener("DOMContentLoaded", async function () {
        digestToken = await getDigest(baseUrl);
        setInterval(async function () {
            digestToken = await getDigest(baseUrl);
        }, 60000);
    });;

    /**
     * @name getListItems
     * @description Função que retorna uma Promise generica que executa o HTTP GET em uma lista do Sharepoint com opção de inserir parametros e retorna um array de objetos
     * @param {string} list Nome da lista (ex.: 'ListaExemploSP')
     * @param {string} path Parametros para filtro, orderby e etc. caso necessário (ex.: '?$select=*,Person/Title,Predio/Title&$expand=Person,Predio')
     * @param {array} arrayResp Parametro para trabalhar com a recursão, normalmente não é necessário preencher.
     */
    this.getListItems = function (list, path, arrayResp) {
        return new Promise(function (resolve, reject) {
            var endUrl = (path ? path : '')
            var xhr = new XMLHttpRequest();
            xhr.open('GET', baseUrl + '/_api/web/lists/GetByTitle(\'' + list + '\')/items' + endUrl, true);
            xhr.responseType = 'json';
            xhr.setRequestHeader('Accept', 'application/json; odata=verbose')
            xhr.onload = async function () {
                if (typeof (xhr.response) != 'object') {
                    var response = JSON.parse(xhr.response)
                } else {
                    var response = xhr.response
                }
                if (arrayResp) {
                    arrayResp = arrayResp.concat(response.d.results)
                    if (response['d']['__next']) {
                        var newPath = response['d']['__next'].substr(Number(response['d']['__next'].toLowerCase().indexOf('items?')) + 5)
                        var resp = await globalScope.getListItems(list, newPath, arrayResp)
                        resolve(resp)
                    } else {
                        resolve(arrayResp)
                    }
                } else {
                    var initialResponseArr = response.d.results
                    if (response['d']['__next']) {
                        var newPath = response['d']['__next'].substr(Number(response['d']['__next'].toLowerCase().indexOf('items?')) + 5)
                        var resp = await globalScope.getListItems(list, newPath, initialResponseArr)
                        resolve(resp)
                    } else {
                        resolve(initialResponseArr)
                    }
                }
            };
            xhr.onerror = function () {
                reject()
            };
            xhr.send();
        });
    }

    /**
     * @name postItemList
     * @description Função que retorna uma Promise que insere em uma lista determinado item passado como um objeto
     * @param {string} list Nome da lista (ex.: 'ListaExemploSP')
     * @param {string} data Objeto a ser inserido na lista, não é necessário __metadata (ex.: { Title: 'Isso é um item' })
     */
    this.postItemList = function (list, data) {
        return new Promise(async function (resolve, reject) {
            if (digestToken.length == 0) {
                digestToken = await getDigest(baseUrl);
            }
            var dataToSend = data
            dataToSend['__metadata'] = {
                type: "SP.Data." + list[0].toUpperCase() + list.substr(1) + "ListItem"
            }
            var xhr = new XMLHttpRequest();
            xhr.open('POST', baseUrl + '/_api/web/lists/GetByTitle(\'' + list + '\')/items', true);
            xhr.responseType = 'json';
            xhr.setRequestHeader('Content-Type', 'application/json;odata=verbose');
            xhr.setRequestHeader('Accept', 'application/json;odata=verbose ');
            xhr.setRequestHeader('X-RequestDigest', digestToken);
            xhr.setRequestHeader('IF-MATCH', '*');
            xhr.onload = function (x) {
                if (typeof (xhr.response) != 'object') {
                    var response = JSON.parse(xhr.response)
                } else {
                    var response = xhr.response
                }
                if (xhr.status > 299) {
                    reject(response)
                }
                resolve(response['d'])
            };
            xhr.onerror = function (e) {
                reject(e)
            };
            xhr.send(JSON.stringify(dataToSend));
        });
    }

    /**
     * @name updateItemList
     * @description Função que retorna uma Promise para atualizar um item de uma lista.
     * @param {string} list Nome da lista (ex.: 'ListaExemploSP')
     * @param {string} id Identificador número do item na lista  (ex.: 53)
     * @param {string} data Objeto a ser substituido na lista, não é necessário __metadata (ex.: { Title: 'Isso é um item' })
     */
    this.updateItemList = function (list, id, data) {
        return new Promise(async function (resolve, reject) {
            if (digestToken.length == 0) {
                digestToken = await getDigest(baseUrl);
            }
            var dataToSend = data
            dataToSend['__metadata'] = {
                type: "SP.Data." + list[0].toUpperCase() + list.substr(1) + "ListItem"
            }
            var xhr = new XMLHttpRequest();
            xhr.open('POST', baseUrl + '/_api/web/lists/GetByTitle(\'' + list + '\')/items(' + id + ')', true);
            xhr.responseType = 'json';
            xhr.setRequestHeader('Content-Type', 'application/json;odata=verbose');
            xhr.setRequestHeader('Accept', 'application/json;odata=verbose');
            xhr.setRequestHeader('X-RequestDigest', digestToken);
            xhr.setRequestHeader('IF-MATCH', '*');
            xhr.setRequestHeader('X-HTTP-Method', 'MERGE');
            xhr.onload = function (x) {
                if (typeof (xhr.response) != 'object') {
                    var response = JSON.parse(xhr.response)
                } else {
                    var response = xhr.response
                }
                if (xhr.status > 299) {
                    reject(x)
                }
                resolve(response)
            };
            xhr.onerror = function () {
                reject()
            };
            xhr.send(JSON.stringify(dataToSend));
        });
    }
    /**
     * @name deleteItemList
     * @description Função que retorna uma Promise para deletar um item em uma lista
     * @param {string} list Nome da lista (ex.: 'ListaExemploSP')
     * @param {string} id Identificador numerico do item na lista  (ex.: 53)
     */
    this.deleteItemList = function (list, id) {
        return new Promise(async function (resolve, reject) {
            if (digestToken.length == 0) {
                digestToken = await getDigest(baseUrl);
            }
            var xhr = new XMLHttpRequest();
            xhr.open('POST', baseUrl + '/_api/web/lists/GetByTitle(\'' + list + '\')/items(' + id + ')/recycle()', true);
            xhr.responseType = 'json';
            xhr.setRequestHeader('Content-Type', 'application/json;odata=verbose');
            xhr.setRequestHeader('Accept', 'application/json;odata=verbose');
            xhr.setRequestHeader('X-RequestDigest', digestToken);
            xhr.setRequestHeader('IF-MATCH', '*');
            xhr.setRequestHeader('X-HTTP-Method', 'DELETE');
            xhr.onload = function () {
                if (typeof (xhr.response) != 'object') {
                    var response = JSON.parse(xhr.response)
                } else {
                    var response = xhr.response
                }
                resolve(response)
            };
            xhr.onerror = function () {
                reject()
            }
            xhr.send();
        });
    }

    /**
     * @name postAttachmentList
     * @description Função que retorna uma Promise que insere um arquivo como anexo em um item de uma lista do Sharepoint
     * @param {string} list Nome da lista (ex.: 'ListaExemploSP')
     * @param {string} id Identificador numerico do item na lista (ex.: 33)
     * @param {element} input Elemento HTML Input do tipo file (ex.: document.getElementById('teste') )
     */
    this.postAttachmentList = function (list, id, input) {
        return new Promise(async function (resolve, reject) {
            if (digestToken.length == 0) {
                digestToken = await getDigest(baseUrl);
            }
            if (input.files.length == 0) {
                throw Error('Input inserido no parametro element não possui um arquivo')
            }
            var fileName = String(input.files[0].name).replace(/[^a-zA-Z0-9. ]/g, ""); // Limitações de caracteres do Sharepoint
            var reader = new FileReader();
            var xhr = new XMLHttpRequest();
            xhr.open('POST', baseUrl + '/_api/web/lists/GetByTitle(\'' + list + '\')/items(' + id + ')/AttachmentFiles/add(FileName=\'' + fileName + '\')', true);
            xhr.responseType = 'json';
            xhr.setRequestHeader('Accept', 'application/json;odata=verbose');
            xhr.setRequestHeader('X-RequestDigest', digestToken);
            xhr.onload = function () {
                if (typeof (xhr.response) != 'object') {
                    var response = JSON.parse(xhr.response)
                } else {
                    var response = xhr.response
                }
                resolve(response['d'])
            };
            xhr.onerror = function () {
                reject()
            };
            reader.onerror = function (e) {
                reject()
            };
            reader.onload = function (e) {
                xhr.send(e.target.result);
            }
            reader.readAsArrayBuffer(input.files[0]);
        });
    }

    /**
     * @name deleteAttachmentList
     * @description Função que retorna uma Promise que deleta um anexo pelo nome de um item na lista do sharepoint 
     * @param {string} list Nome da lista (ex.: 'ListaExemploSP')
     * @param {string} id Identificador numerico do item na lista (ex.: 33)
     * @param {string} filename Nome do arquivo com extensão (ex.: 'fotolegal.jpeg' )
     */
    this.deleteAttachmentList = function (list, id, filename) {
        return new Promise(async function (resolve, reject) {
            if (digestToken.length == 0) {
                digestToken = await getDigest(baseUrl);
            }
            var xhr = new XMLHttpRequest();
            xhr.open('POST', baseUrl + '/_api/web/lists/GetByTitle(\'' + list + '\')/items(' + id + ')/AttachmentFiles/getByFileName(\'' + encodeURIComponent(filename) + '\')', true);
            xhr.responseType = 'json';
            xhr.setRequestHeader('Accept', 'application/json;odata=verbose');
            xhr.setRequestHeader('X-RequestDigest', digestToken);
            xhr.setRequestHeader('IF-MATCH', '*');
            xhr.setRequestHeader('X-HTTP-Method', 'DELETE');
            xhr.onload = function () {
                if (typeof (xhr.response) != 'object') {
                    var response = JSON.parse(xhr.response)
                } else {
                    var response = xhr.response
                }
                resolve(response)
            };
            xhr.onerror = function () {
                reject()
            };
            xhr.send();
        });
    }

    /**
     * @name getCurrentUser
     * @description Função que retorna uma Promise que busca as informações do usuário que está acessando a página
     */
    this.getCurrentUser = function () {
        return new Promise(function (resolve, reject) {
            var xhr = new XMLHttpRequest();
            xhr.open('GET', baseUrl + '/_api/web/currentuser', true);
            xhr.responseType = 'json';
            xhr.setRequestHeader('Accept', 'application/json; odata=verbose')
            xhr.onload = function () {
                if (typeof (xhr.response) != 'object') {
                    var response = JSON.parse(xhr.response)
                } else {
                    var response = xhr.response
                }
                resolve(response['d'])
            };
            xhr.onerror = function () {
                reject()
            };
            xhr.send();
        });
    }

    /**
     * @name getCurrentUserMore
     * @description Função que retorna uma Promise que busca as informações do usuário que está acessando a página como a função anterior, porém com algumas informações a mais como Departamento, Manager, Cellphone...
     */
    this.getCurrentUserMore = function () {
        return new Promise(function (resolve, reject) {
            var xhr = new XMLHttpRequest();
            xhr.open('GET', baseUrl + '/_api/SP.UserProfiles.PeopleManager/GetMyProperties?$select=*', true);
            xhr.responseType = 'json';
            xhr.setRequestHeader('Accept', 'application/json; odata=verbose')
            xhr.onload = function () {
                if (typeof (xhr.response) != 'object') {
                    var response = JSON.parse(xhr.response)
                } else {
                    var response = xhr.response
                }
                resolve(response['d'])
            };
            xhr.onerror = function () {
                reject()
            };
            xhr.send();
        });
    }

    /**
     * @name getSpecifyUserMore
     * @description  Função que retorna uma Promise que busca as informações de um usuário específico do site
     * @param {string} userName Username do usuário (ex.: 'i:0#.f|membership|nome.sobrenome@obuc.com')
     */
    this.getSpecifyUserMore = function (userName) {
        return new Promise(function (resolve, reject) {
            var xhr = new XMLHttpRequest();
            xhr.open('GET', baseUrl + "/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='" +
                encodeURIComponent(userName) + "'", true);
            xhr.responseType = 'json';
            xhr.setRequestHeader('Accept', 'application/json; odata=verbose')
            xhr.onload = function () {
                if (typeof (xhr.response) != 'object') {
                    var response = JSON.parse(xhr.response)
                } else {
                    var response = xhr.response
                }
                resolve(response['d'])
            };
            xhr.onerror = function () {
                reject()
            };
            xhr.send();
        });
    }
    /**
     * @name getGroups
     * @description Função que retorna uma Promise que busca todos os grupos deste Site de Sharepoint 
     */
    this.getGroups = function () {
        return new Promise(function (resolve, reject) {
            var xhr = new XMLHttpRequest();
            xhr.responseType = 'json';
            xhr.open('GET', baseUrl + '/_api/web/sitegroups', true);
            xhr.setRequestHeader('Accept', 'application/json; odata=verbose')
            xhr.onload = function () {
                if (typeof (xhr.response) != 'object') {
                    var response = JSON.parse(xhr.response)
                } else {
                    var response = xhr.response
                }
                resolve(response['d']['results'])
            };
            xhr.onerror = function () {
                reject()
            };
            xhr.send();
        });
    }


    /**
     * @name getUsersGroup
     * @description Função que retorna uma Promise que busca os usuários de determinado grupo
     * @param {string} idgroup Id do grupo (ex.: 4)
     */
    this.getUsersGroup = function (idgroup) {
        return new Promise(function (resolve, reject) {
            var xhr = new XMLHttpRequest();
            xhr.open('GET', baseUrl + '/_api/web/sitegroups/getbyid(' + idgroup + ')/users', true);
            xhr.responseType = 'json';
            xhr.setRequestHeader('Accept', 'application/json; odata=verbose')
            xhr.onload = function () {
                if (typeof (xhr.response) != 'object') {
                    var response = JSON.parse(xhr.response)
                } else {
                    var response = xhr.response
                }
                resolve(response['d']['results'])
            };
            xhr.onerror = function () {
                reject()
            };
            xhr.send();
        });
    }


    /**
     * @name postUserToGroup
     * @description Função que retorna uma Promise que adiciona usuário no grupo desejado
     * @param {string} LoginName Identificador de Login completo (ex.: 'i:0#.f|membership|dostoievski.batista@obuc.com.br')
     * @param {string} idgroup Id do grupo aonde a pessoa será adicionada (ex.: 4)
     */
    this.postUserToGroup = function (loginName, idgroup) {
        return new Promise(async function (resolve, reject) {
            if (digestToken.length == 0) {
                digestToken = await getDigest(baseUrl);
            }
            var dataToSend = {
                __metadata: {
                    type: "SP.User"
                },
                LoginName: loginName
            }
            var xhr = new XMLHttpRequest();
            xhr.open('POST', baseUrl + '/_api/web/sitegroups/GetById(' + idgroup + ')/users', true);
            xhr.responseType = 'json';
            xhr.setRequestHeader('Content-Type', 'application/json;odata=verbose');
            xhr.setRequestHeader('Accept', 'application/json;odata=verbose');
            xhr.setRequestHeader('X-RequestDigest', digestToken);
            xhr.onload = function () {
                if (typeof (xhr.response) != 'object') {
                    var response = JSON.parse(xhr.response)
                } else {
                    var response = xhr.response
                }
                resolve(response['d']['results'])
            };
            xhr.onerror = function () {
                reject()
            };
            xhr.send(JSON.stringify(dataToSend));
        });
    }
    /**
     * @name SendEmail
     * @description Função para enviar email para emails de outros usuarios (que tem o acesso ao Site).
     * @param {array} tolist Lista de pessoas que vão receber o Email  (ex.: ['ususario1.sobrenome@obuc.com.br','ususario1.sobrenome@obuc.com.br'])
     * @param {array} cclist Lista de pessoas que vão receber o Email em cópia  (ex.: ['ususario1.sobrenome@obuc.com.br','ususario1.sobrenome@obuc.com.br'])
     * @param {string} subject Assunto do Email (ex.:'Atualização do projeto' )
     * @param {string} body Body da mensagem (ex.: 'Ocorreu uma atualização no dia 22/12/2022')
     */
    this.sendEmail = function (tolist, cclist, subject, body) {
        return new Promise(async function (resolve, reject) {
            if (digestToken.length == 0) {
                digestToken = await getDigest(baseUrl);
            }
            var dataToSend = {
                'properties': {
                    'To': {
                        'results': tolist
                    },
                    'CC': {
                        'results': cclist
                    },
                    'Subject': subject,
                    'Body': body,
                    "AdditionalHeaders": {
                        "__metadata": {
                            "type": "Collection(SP.KeyValue)"
                        },
                        "results": [{
                            "__metadata": {
                                "type": 'SP.KeyValue'
                            },
                            "Key": "content-type",
                            "Value": 'text/html',
                            "ValueType": "Edm.String"
                        }]
                    }
                }
            };
            dataToSend['properties']['__metadata'] = {
                type: "SP.Utilities.EmailProperties"
            }
            var xhr = new XMLHttpRequest();
            xhr.open('POST', baseUrl + '/_api/SP.Utilities.Utility.SendEmail', true);
            xhr.responseType = 'json';
            xhr.setRequestHeader('Content-Type', 'application/json;odata=verbose');
            xhr.setRequestHeader('Accept', 'application/json;odata=verbose');
            xhr.setRequestHeader('X-RequestDigest', digestToken);
            xhr.setRequestHeader('contentType', 'application/json');

            xhr.onload = function (x) {
                if (typeof (xhr.response) != 'object') {
                    var response = JSON.parse(xhr.response)
                } else {
                    var response = xhr.response
                }
                if (xhr.status > 299) {
                    reject(x)
                }
                resolve(response)
            };
            xhr.onerror = function () {
                reject()
            };
            xhr.send(JSON.stringify(dataToSend));
        });
    }

    /**
     * @name getDigest
     * @description Função que executa a busca do token, necessário para execução de HTTP request (exceção do metodo GET)
     * @param {string} baseUrl Endereço do Sharepoint Site (ex.: https://obuc.sharepoint.com/sites/Development)
     */
    function getDigest(baseUrl) {
        return new Promise(function (resolve, reject) {
            var xhr = new XMLHttpRequest();
            xhr.open('POST', baseUrl + '/_api/contextinfo', true);
            xhr.responseType = 'json';
            xhr.setRequestHeader('Accept', 'application/json; odata=verbose')
            xhr.onload = function () {
                if (typeof (xhr.response) != 'object') {
                    var response = JSON.parse(xhr.response)
                } else {
                    var response = xhr.response
                }
                if (this.status >= 200 && this.status < 300) {
                    var digestT = response['d']['GetContextWebInformation']['FormDigestValue']
                    resolve(digestT)
                } else {
                    reject()
                }
            };
            xhr.onerror = function () {
                reject()
            };
            xhr.send();
        })
    }

}
export default new CrudSharepoint(url)