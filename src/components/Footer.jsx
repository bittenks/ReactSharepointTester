import React, { Component } from "react";

export default class Footer extends Component {
  render() {
    let endpoint =`https://source.unsplash.com/random/2400x3000/?gradient`
    const getImage = async (endpoint) => {
      return await fetch(endpoint).then(res => res.url)
    }
    getImage(endpoint)
    .then(result => {
      setLoading(false);
    })
    return (
      <footer className="p-4 bg-white rounded-lg shadow md:flex md:items-center md:justify-between md:p-6 dark:bg-gray-800">
        <span className="text-sm text-gray-500 sm:text-center dark:text-gray-400">
          © 2022 Card Alpha Tester React + Vite
        </span>
        <button
              type="button"
              className="text-sm border-white  cursor-pointer mt-2 mb-2  text-white"
             onClick={() => getImage('https://source.unsplash.com/random/2400x3000/?gradient').then(result => {
              document.body.style.background = ` url('${result}') no-repeat `
             ;
            })}
            >
              🎨 Random Background
            </button>
            <a
              href="https://github.com/bittenks/ReactSharepointTester"
              className="mr-4 hover:underline md:mr-6 "
            >
              Github
            </a>
      </footer>
    );
  }
}
