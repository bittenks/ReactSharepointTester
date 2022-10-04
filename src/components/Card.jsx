import React, { useState, useEffect } from "react";
import { ToastContainer, toast } from "react-toastify";

import "react-toastify/dist/ReactToastify.css";
import "./card.css";
import $ from "jquery";
import sp from "../services/sp.full";
import createPeoplePicker from "../services/PeoplePicker";
import Loading from "./Loading";

export default function Card() {
  const [loading, setLoading] = useState(true);
  const search = window.location.search;
  const params = new URLSearchParams(search);
  const foo = params.get("user");
  const [user, setUser] = useState({});
  const [name, setName] = useState("");
  const [userD, setUserD] = useState({
    Bio: "",
    Cellphone: "",
    Skills: "",
  });
  const notify = () => {
    toast.warn("🦄 Wow so easy!", {
      position: "top-center",
      autoClose: 5000,
      hideProgressBar: false,
      closeOnClick: true,
      pauseOnHover: true,
      draggable: true,
      progress: undefined,
    });
  };
  const GerarCard = () => {
    if ($("#username").attr("emailuser")) {
      window.open(
        "https://bayergroup.sharepoint.com/sites/022971/cardAlpha/index.html?user=" +
          $("#username").attr("emailuser")
      );
    } else {
      alert("Precisa inserir um usuario válido");
    }
  };

  useEffect(() => {
    if (foo) {
      const timer = setTimeout(() => {
        const fetchCurrentUser = async () => {
          const currentUserFromServer = await sp.getSpecifyUserMore(
            "i:0#.f|membership|" + foo
          );
          const ResultsUser =
            currentUserFromServer.UserProfileProperties.results;
          var BioUser = ResultsUser.find((user) => {
            return user.Key == "AboutMe";
          });
          var SkillsUser = ResultsUser.find((user) => {
            return user.Key == "SPS-Skills";
          });
          var CellUser = ResultsUser.find((user) => {
            return user.Key == "CellPhone";
          });
          const CurrentuserD = {
            Bio: BioUser.Value,
            Cellphone: CellUser.Value,
            Skills: SkillsUser.Value,
          };
          console.log(ResultsUser);
          console.log(BioUser.Value);

          setUser(currentUserFromServer);
          setUserD(CurrentuserD);
        };
        fetchCurrentUser();
        setLoading(false);
      }, 3000);
      return () => clearTimeout(timer);
    } else {
      const timer = setTimeout(() => {
        const fetchCurrentUser = async () => {
          const currentUserFromServer = await sp.getCurrentUserMore();
          setUser(currentUserFromServer);
        };
        fetchCurrentUser();
        console.log(name);
        setLoading(false);
      }, 2000);
      return () => clearTimeout(timer);
    }
  }, []);
  const ImageUser =
    "https://bayergroup.sharepoint.com/sites/020693/_layouts/15/userphoto.aspx?size=L&username=" +
    user.Email;
  const MailSend = "mailto:" + user.Email;
  const LinkUser =
    "https://bayergroup.sharepoint.com/sites/022971/cardAlpha/index.html?user=" +
    user.Email;
  if (foo) {
    return (
      <div className="bg-gradient-to-r from-cyan-500 to-blue-500 font-sans h-screen w-full flex flex-row justify-center items-center">
        {loading && <Loading />}
        {!loading && (
          <div className="card cardPrincipal w-96 mx-auto bg-white rounded-lg  shadow-xl hover:shadow">
            <img
              src={ImageUser}
              alt="foto de perfil"
              className="w-32 mx-auto rounded-full -mt-20 border-2 border-white"
            />
            <div className="text-center mt-2 text-3xl font-medium">
              {user.DisplayName}
            </div>
            <div className="text-center mt-2 font-light text-sm">
              @{user.Title}
            </div>
            <div className="text-center mt-2 font-normal text-lg">
              <a href={MailSend} className="btn-primary">
                Contact 📩
              </a>
            </div>

            <div className="px-6 text-center mt-2 font-light text-sm overflow-y-auto h-20">
              <p className="text-sm">{userD.Bio}</p>
            </div>
            <hr className="mt-8" />
            <div className="flex p-4">
              <div className="w-3/4 overflow-x-hidden text-center overflow-y-auto h-12 ">
                <span className="font-bold">Skills ⚡</span>
                <p className="text-sm">{userD.Skills}</p>
              </div>
              <div className="w-0 border border-gray-300"></div>
              <div className="w-2/3 overflow-x-hidden text-center overflow-y-auto h-12">
                <span className="font-bold">Cellphone 📞</span>
                <p className="text-sm">{userD.Cellphone}</p>
              </div>
            </div>
          </div>
        )}
      </div>
    );
  } else {
    return (
      <div className="bg-gradient-to-r from-cyan-500 to-blue-500 font-sans h-screen w-full flex flex-row justify-center items-center">
        <ToastContainer />
        {loading && <Loading />}
        {!loading && (
          <div
            onLoad={createPeoplePicker(
              "https://bayergroup.sharepoint.com/sites/022971/"
            )}
            className="card w-60 text-center bg-gray-50 rounded-lg  shadow-xl hover:shadow"
          >
            <a href={LinkUser} className="mt-2  btn-primary m-1 ">
              👋 Gerar Seu Card
            </a>
            <label
              className="mt-2 block text-gray-700 text-sm font-bold mb-2"
              for="username"
            >
              Username :
            </label>
            <input
              className="singlePeople  mt-2 shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline bg-transparent"
              id="username"
              type="text"
              placeholder="Buscar usuario"
              onChange={(e) => {
                setName(e.target.getAttribute("emailuser"));
              }}
              onInput={(e) => {
                setName(e.target.getAttribute("emailuser"));
              }}
            />
            <button
              type="button"
              className="cursor-pointer mt-2 mb-2  text-slate-900"
              onClick={GerarCard}
            >
              🔎 Gerar User Card
            </button>
          </div>
        )}
      </div>
    );
  }
}
