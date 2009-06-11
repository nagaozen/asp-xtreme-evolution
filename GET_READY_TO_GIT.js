/* +-----------------------------------------------------------------------------+
 * |This file is part of ASP Xtreme Evolution.                                   |
 * |Copyright © 2007, 2008 Fabio Zendhi Nagao                                    |
 * |                                                                             |
 * |ASP Xtreme Evolution is free software: you can redistribute it and/or modify |
 * |it under the terms of the GNU Lesser General Public License as published by  |
 * |the Free Software Foundation, either version 3 of the License, or            |
 * |(at your option) any later version.                                          |
 * |                                                                             |
 * |ASP Xtreme Evolution is distributed in the hope that it will be useful,      |
 * |but WITHOUT ANY WARRANTY; without even the implied warranty of               |
 * |MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the                |
 * |GNU Lesser General Public License for more details.                          |
 * |                                                                             |
 * |You should have received a copy of the GNU Lesser General Public License     |
 * |along with ASP Xtreme Evolution. If not, see <http://www.gnu.org/licenses/>. |
 * +-----------------------------------------------------------------------------+
 * 
 * File: UPDATE_README.js
 * 
 * This file is not part of the framework. The author uses it to update the 
 * github README.md
 * 
 * About:
 * 
 *   - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ June 2009
 */

try {
    var readme = [
    "ASP Xtreme Evolution",
    "====================",
    "",
    "The ASP Xtreme Evolution goal is to be a versatile MVC URL-Friendly base for Classic ASP applications with some additional features that are not ASP native. It should implement things that are common to most applications removing the pain of starting a new software and helping you to structure it so that you get things right from the beginning.",
    "",
    "License",
    "-------"
    ].join('\n');
    var Fso = new ActiveXObject("Scripting.FileSystemObject");
    var File = null;
    
    // Append License
    File = Fso.openTextFile("app/docs/LICENSE");
    readme += "\n" + File.readAll();
    File.close();
    
    // Append Installation
    File = Fso.openTextFile("app/docs/INSTALL.md");
    readme += "\n" + File.readAll();
    File.close();
    
    // Append Changes
    File = Fso.openTextFile("app/docs/CHANGES.md");
    readme += "\n" + File.readAll();
    File.close();
    
    // Save it
    File = Fso.createTextFile("README.md", true, false);
    File.write(readme);
    File.close();
    
    WScript.echo("README.md generation is ... completed");
    
    File = null;
    Fso = null;
    
    // Clean the cache directory
    var Shell = new ActiveXObject("WScript.Shell");
    Shell.run("cmd /c del /F/Q app\\cache\\*.*");
    Shell = null;
    
    WScript.echo("app/cache cleaning is ..... completed");
} catch(e) {
    WScript.echo(e);
} finally {
    WScript.echo("AXE is ready to be pushed into github");
}
