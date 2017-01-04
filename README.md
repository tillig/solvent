# Solvent - Visual Studio Extension for Solution Explorer

[Back in 2004](http://www.paraesthesia.com/archive/2004/06/25/solvent-power-toys-for-visual-studio-.net.aspx/) I entered a VS 2003 extension contest [and won second place](http://osherove.com/blog/2004/8/13/add-in-contest-winners-announced.html) with this extension.

Solvent added the following functions to the Visual Studio Solution Explorer:

- Recursively expand/contract items (expand all/collapse all)
- Select an item or project and open the containing folder in Windows Explorer
- Open a command prompt in the same folder as an item or a project

I never took it past VS 2003 since, over time, these features actually ended up being baked into Visual Studio. However, it remains an example of how to make a VS extension and connect to VS using the DTE structure (which, as I recall, was a huge pain).