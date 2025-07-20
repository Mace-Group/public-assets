# public-assets
Simple static resources used by the digital development team solutions

# How to Reference the Modules
Many of the files here are modules such as JavaScript files that may need to be referenced by pages using *Script* Tags 

```html
<script src="some/path/to/resource.js"></script>
```

The problem is you do not necessarily want to just copy the file from this repository and host on your own site

(You might, every use case has it's requirements!)

To simply use the files you find here you *might* think that the RAW url you use to see the code is teh way forward, the problem is that the GitHub server will send out the file with a Mime Type of _text/plain_ which will generally be blocked from being interpreted as a JavaScript resource!

What you need to do is use a CDN network that gives the correct type. One such CDN is _jsdelivr_

https://cdn.jsdelivr.net/gh/Mace-Group/public-assets/
