<div id="top"></div>
<!--

-->





<!-- PROJECT LOGO -->
<br/>
<div align="center">
  <a href="https://github.com/asadzz/Powerpoint-VBscript-hack/blob/main/">
    <img src="images/ppt%20hack.jpg" alt="Logo" width="80" height="80">
  </a>

  <h3 align="center">Readme</h3>

  <p align="center">
        Powerpoint VBscript Hack!
    <br/>
    <a href="https://github.com/asadzz/Powerpoint-VBscript-hack"><strong>Explore the docs »</strong></a>
   <br/>
    <a href="https://github.com/asadzz/Powerpoint-VBscript-hack/issues">Report Bug</a>
     
    
   
  </p>
</div>



<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#about-the-project">About The Project</a>
      <ul>
        <li><a href="#built-with">Built With</a></li>
      </ul>
    </li>
    <li>
      <a href="#getting-started">Getting Started</a>
      <ul>
        <li><a href="#prerequisites">Prerequisites</a></li>
        <li><a href="#installation">Installation</a></li>
      </ul>
    </li>
    <li><a href="#usage">Usage</a></li>
    <li><a href="#contributing">Contributing</a></li>
    <li><a href="#license">License</a></li>
    <li><a href="#contact">Contact</a></li>
    <li><a href="#acknowledgments">Acknowledgments</a></li>
  </ol>
</details>



<!-- ABOUT THE PROJECT -->
## About The Project

[![Product Name Screen Shot](https://github.com/asadzz/Powerpoint-VBscript-hack/blob/main/images/ppt%20hack.jpg)

This simple yet powerful VBscript method can be used to progamatically change layout and content of slides. Do give it a try. I have put comments in code for easy undestanding


Use the `README.md` to get started.

<p align="right">(<a href="#top">back to top</a>)</p>



### Built With

This section should list any major frameworks/libraries used to bootstrap your project. Leave any add-ons/plugins for the acknowledgements section. Here are a few examples.

* To run this code, simply open Developer tab in PowerPoint Depending upon the version of PowerPoint you need to enable this tab using "**customize ribbon**" option under "**PowerPoint options**",
* Place the code change the function to one you want or keep as it Run the macro and see changes as per logic defined. I have put comments in code for understanding and changing.

<p align="right">(<a href="#top">back to top</a>)</p>


<!-- GETTING STARTED -->
## Getting Started

The best place to start is to explore VB functions in MS-powerpoint using MSDN [library](https://docs.microsoft.com/en-us/office/vba/api/overview/powerpoint)

### Prerequisites

* Powerpoint with Macro-enabled


<!-- USAGE EXAMPLES -->
## Usage


``  oSl.Shapes(1).TextFrame.TextRange.Paragraphs(1).Lines(1, 1).Text = "" ``
where ``os1`` is object of ``ActivePresentation.Slides``. 

![image](https://user-images.githubusercontent.com/7777434/163054864-04800ac0-c6dc-4bc3-a6f1-ee21bc818cf0.png)

shape(1) will select ``main-heading`` and ``paragraph`` represent orders (1,2)and ``Lines`` be donate lines of paragraph. The above line will thus delete the **main-heading**. The ``paragraph`` is more relevant in text-box if you want to change the first-para heading you will select "paragraph(2).".

Please see the code for further comments.
_For more examples, please refer to the [Documentation](https://docs.microsoft.com/en-us/office/vba/api/overview/powerpoint)_

<p align="right">(<a href="#top">back to top</a>)</p>




<!-- CONTRIBUTING -->
## Contributing

Contributions are what make the open source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement".
Don't forget to give the project a star! Thanks again!

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

<p align="right">(<a href="#top">back to top</a>)</p>



<!-- LICENSE -->
## License

Distributed under the MIT License. See `LICENSE.txt` for more information.

<p align="right">(<a href="#top">back to top</a>)</p>



<!-- CONTACT -->
## Contact

Your Name - [@Catchabyte1](https://twitter.com/Catchabyte1) - a.alii85@gmail.com

Project Link: [https://github.com/asadzz/Powerpoint-VBscript-hack](https://github.com/asadzz/Powerpoint-VBscript-hack)

<p align="right">(<a href="#top">back to top</a>)</p>



<!-- ACKNOWLEDGMENTS -->
## Acknowledgments

Use this space to list resources you find helpful and would like to give credit to. I've included a few of my favorites to kick things off!

* [Choose an Open Source License](https://choosealicense.com)


<p align="right">(<a href="#top">back to top</a>)</p>



