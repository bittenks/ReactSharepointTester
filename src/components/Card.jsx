import React, { useState, useEffect } from "react";
import "./card.css";
import sp from "../services/sp.full";

export default function Card() {
  const search = window.location.search;
  const params = new URLSearchParams(search);
  const foo = params.get("user");
  const [user, setUser] = useState({});
  const [userD, setUserD] = useState({
    Bio: "",
    Cellphone: "",
    Skills: "",
  });
  useEffect(() => {
    if (foo) {
      const fetchCurrentUser = async () => {
        const currentUserFromServer = await sp.getSpecifyUserMore(
          "i:0#.f|membership|" + foo
        );
        const ResultsUser = currentUserFromServer.UserProfileProperties.results;
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
    } else {
      const fetchCurrentUser = async () => {
        const currentUserFromServer = await sp.getCurrentUserMore();
        setUser(currentUserFromServer);
      };
      fetchCurrentUser();
    }
  }, []);
  console.log(user);
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
      </div>
    );
  } else {
    return (
      <div className="bg-gradient-to-r from-cyan-500 to-blue-500 font-sans h-screen w-full flex flex-row justify-center items-center">
        <div className="card w-36 text-center bg-gray-50 rounded-lg  shadow-xl hover:shadow">
          <a href={LinkUser} className="btn-primary m-1 ">
            Gerar Card
          </a>
        </div>
      </div>
    );
  }
}
