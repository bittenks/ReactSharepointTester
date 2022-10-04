import $ from "jquery";
var UsersObjects = [];
var UsersObjectsInsert = [];
const isProduction = process.env.NODE_ENV !== "development"
const url = isProduction ? "https://bayergroup.sharepoint.com/sites/022971/" : "http://localhost:8080/"

/**
 * @name PeoplePicker
 * @description Função para ativar automaticamente campos de single e multi select de Users do Sharepoint. : https://github.com/Obuc/PeoplePicker
 * @param {string} baseUrl Endereço do Sharepoint Site (ex.: https://obuc.sharepoint.com/sites/Development)
 */
export default function createPeoplePicker(baseUrl) {

    // Ativa todas as <divs> com classe .multiPeople para virar campo de multiselect people
    var countUL = 0
    $(".multiPeople").each(function () {
        $(this).attr("ppId", countUL)
        $(this).addClass("form-control")
        $(this).append("<input class='classPPUser multiPPInput' style='border: none; outline:none' ppId='" + countUL + "'/>")
        countUL++
    })

    // Ativa todos os <inputs> com classe .singlePeople para virar campo de people
    $(".singlePeople").each(function () {
        console.info('Eu identifiquei')
        $(this).addClass("classPPUser")
    })

    //Lista do search sendo criada
    var lista = "<ul id = 'listaUsuario' style='position: absolute; z-index: 10099;  list-style: none; max-height: 17vw; overflow-y: auto; color: black; padding-left:10px; background-color:white;display: none'></ul>"
    $('body').append(lista)

    // posicionamento da lista e populando de acordo com valor buscado
    $(".classPPUser").off().on('input', function () {
        var inputId = $(this).attr('id')
        var mb = $(this).css("margin-bottom").substring(0, $(this).css("margin-bottom").length - 2)
        var mt = $(this).css("margin-top").substring(0, $(this).css("margin-top").length - 2)
        var pb = $(this).css("padding-bottom").substring(0, $(this).css("padding-bottom").length - 2)
        var pt = $(this).css("padding-top").substring(0, $(this).css("padding-top").length - 2)
        var spacoTotal = Number(mb) + Number(mt) + Number(pb) + Number(pt)
        $("#listaUsuario").css("top", ($(this).offset().top + $(this).height()) + spacoTotal);
        $("#listaUsuario").css("left", $(this).offset().left)
        $("#listaUsuario").css("width", $(this).outerWidth())
        $("#listaUsuario").empty()
        $("#listaUsuario").show()
        var a = 0
        if ($(this).hasClass("multiPPInput")) {
            var tipocolocarUser = "multi"
            inputId = $(this).attr('ppId')
        } else {
            var tipocolocarUser = "single"
        }
        for (var i = 0; i < UsersObjects.length; i++) {
            if (UsersObjects[i].Title.toUpperCase().indexOf($(this).val().toUpperCase()) != -1) {
                $("#listaUsuario").append("<li class='UserPeople' style='cursor: pointer' loginName='" + UsersObjects[i].LoginName + "' inputWhere='" + inputId + "')'>" + " <span data-toggle='tooltiptext1' title=" + UsersObjects[i].Email + ">" + UsersObjects[i].Title + "</span> </li>")
                a++
            }
            if (a > 30) {
                break
            }
        }


        $("#listaUsuario li").off().on("click", function () {
            if (tipocolocarUser == "multi") {
                colocarUserDiv($(this).attr("loginName"), $(this).attr("inputWhere"))
            } else {
                colocarUserGeral($(this).attr("loginName"), $(this).attr("inputWhere"))
            }
        })
        Check(a, UsersObjects.length)
    });

    $(".classPPUser").on('focusout', function () {
        setTimeout(function () {
            $("#listaUsuario").hide()
        }, 300)
    });

    $(".multiPeople").on("click", function () {
        $(this).find(".classPPUser").focus()
    })
    /**
     * Função para adicionar Loding e fazer a verificacao se existe usuarios antes e depois de carregar todos os usuarios do Sharepoint
     * @param {number} a Quantidade de Usuarios de acordo com input
     * @param {Users} Users Quantidade de Users do Sharepoint
     */
    function Check(a, Users) {
        $('.UserPeople').hide()
        $("#listaUsuario").append("<li id='LoadUser' style='cursor: pointer'>" + " <span class='loadingSpan' data-toggle='tooltiptext1' title=Loading...>Loading   <div class='spinner-3' ></div></span> </li>")
        setTimeout(function () {
            $('.UserPeople').show()
            if (a == 0) {
                if (Users < 10000) {
                    $("#listaUsuario").empty()
                    $("#listaUsuario").append("<li id='NotFoundUser'>" + " <span class='loadingSpan' data-toggle='tooltiptext1' title=Loading... >loading users...  <div class='spinner-3' ></div> </span> </li>")
                }
                if (Users > 10000) {
                    $("#listaUsuario").empty()
                    $("#listaUsuario").append("<li id='NotFoundUser'>" + " <span data-toggle='tooltiptext1' title=Loading... class='loadingSpan' >User Not Found ☹</span> </li>")
                }
            }
            if (a > 0) {
                $("#NotFoundUser").hide()
                $("#LoadUser").hide()
            }
        }, 500)
    }
    /**
     * Função para inserir valor selecionado para campo de multi people
     * @param {string} LoginName do usuario 
     * @param {string} input Referencia de qual campo vai inserir a informação
     */
    function colocarUserDiv(name, input) {
        $("#listaUsuario" + input).hide()
        var objUsuario = [];
        objUsuario = GetUserDados(name, objUsuario)
        if (objUsuario != -1) {
            $(".multiPeople").each(function () {
                if ($(this).attr('ppId') == input) {
                    nome = objUsuario.Title
                    strUserDiv = "<span class='divPeople selectpeopple' nomeUser='" + objUsuario.Title + "' idUser='" + objUsuario.Id + "'><img class='divPeopleImg' src=\"" + baseUrl + "/_layouts/15/userphoto.aspx?size=L&username=" + objUsuario.UserPrincipalName + "\">" + "<span data-toggle='tooltiptext1' title=" + objUsuario.Email + ">" + nome + "</span><i style='cursor:pointer' class='retirarUser'> x </i></span>"
                    $(strUserDiv).insertBefore($(this).find(".classPPUser"))
                    $(this).find(".classPPUser").val('')
                    $(this).find(".classPPUser").focus()
                }
            })
        }

        $(".retirarUser").off().on("click", function () {
            $(this).parent().remove()
        })
    }

    /**
     * Função para inserir valor selecionado para campo de single people
     * @param {string} LoginName do usuario 
     * @param {string} input Referencia de qual campo vai inserir a informação
     */

    function colocarUserGeral(name, input) {
        $("#listaUsuario").hide()
        var objUsuario = [];
        objUsuario = GetUserDados(name, objUsuario)
        if (objUsuario != -1) {
            $("#" + input).val(objUsuario.Title)
            $("#" + input).attr("idUser", objUsuario.Id)
            $("#" + input).attr("EmailUser", objUsuario.Email)

        }
    }

    //Função para buscar usuarios registrados no SPO referenciado
    function LoadUsarios() {
        $.ajax({
            url: baseUrl + "/_api/web/siteusers",
            type: "GET",
            headers: {
                "accept": "application/json;odata=verbose",
            },
            success: function (data) {
                UsersObjects = data.d.results;
                UsersObjectsInsert = data.d.results;
                // console.log(UsersObjects);
                LoadAllUsers();
            },
            error: function (error) {
                // console.log(JSON.stringify(error));
            }
        });
    }

    //Função para buscar todos os usuarios do SPO mãe
    function LoadAllUsers() {
        $.ajax({
            url: baseUrl.substring(0, baseUrl.indexOf('.com') + 4) + "/_api/web/siteusers",
            type: "GET",
            headers: {
                "accept": "application/json;odata=verbose",
            },
            success: function (data) {
                UsersObjects = data.d.results;
                console.log(UsersObjects)
            },
            error: function (error) {
                //alert(JSON.stringify(error));
            }
        });
    }

    LoadUsarios()

    //Função para pegar os dados do usuario selecionado
    function GetUserDados(userName, varId) {

        var accountName = userName;

        var varAjax = $.ajax({
            url: baseUrl + "/_api/web/siteusers(@v)?@v='" +
                encodeURIComponent(accountName) + "'",
            method: "GET",
            async: false,
            headers: {
                "Accept": "application/json; odata=verbose"
            },
            success: function (data) {
                ///user id received from site users.
                varId = data.d
            },
            error: function (data) {
                varId = postNewUser(userName, varId)
            }
        })
        return varId

    }


    // Função para adicionar usuario no SPO caso ele não esteja registrado no SPO referenciado
    function postNewUser(name, varId) {
        $.ajax({
            url: baseUrl + "/_api/web/ensureuser",
            type: "POST",
            contentType: "application/json;odata=verbose",
            data: "{ 'logonName': '" + name + "' }",
            async: false,
            headers: {
                "X-RequestDigest": digestToken, // digestToken sendo gerado no crudSharepoint
                "accept": "application/json;odata=verbose"
            },
            success: function (data) {

                varId = data.d
            },
            error: function (data) {
                alert(JSON.stringify(data.responseJSON.error.message.value));
                varId = -1
            }
        })
        return varId
    }
}