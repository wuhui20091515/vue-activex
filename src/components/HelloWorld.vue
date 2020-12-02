<template>
  <div class="hello">
    <button @click="OpenDoc">Show/Hide Titlebar </button>
    <div id="oframediv">
    </div>
    <div v-html="message">
    </div>
  </div>
</template>

<script>
import script from "@/components/script";
export default {
  name: 'HelloWorld',
  props: {
    msg: String
  },
  data:function() {
  return {
    ocx:'',
    param:'',
    oframex:'',
    message:"<h1>爱你哦！</h1>"
  };
  },
  mounted() {
    let BorderStyle = document.createElement('param')
    BorderStyle.setAttribute('name','BorderStyle')
    BorderStyle.setAttribute('value','50')
    let TitlebarColor = document.createElement('param')
    TitlebarColor.setAttribute('name',"TitlebarColor")
    TitlebarColor.setAttribute('value', '52479')

    let scriptStr = document.createElement('object')
    scriptStr.setAttribute('id','oframe')
    scriptStr.setAttribute('classid','clsid:00460182-9E5E-11d5-B7C8-B8269041DD57')
    scriptStr.setAttribute('height',"800")
    scriptStr.setAttribute('width',"800")
    scriptStr.setAttribute('codebase','ActiveX/DSOframer/DSOframer.CAB#version=1,0,0,0')
    scriptStr.appendChild(BorderStyle)
    scriptStr.appendChild(TitlebarColor)
    this.ocx=document.getElementById("oframediv")
    this.ocx.appendChild(scriptStr)
    this.initActiveXObject()
    let x=document.getElementById('oframediv')
    script.oframe=x.childNodes.item(0)
   // this.initPlugin()
  },
  methods:{
    initPlugin(){
      console.log('plugin', this.$refs.oframe)
    },
    initActiveXObject () {
      let scriptStrs = document.createElement('script')
      scriptStrs.setAttribute('type','text/javascript')
      scriptStrs.setAttribute('language','jscript')
      scriptStrs.setAttribute('for', 'oframe')
      scriptStrs.setAttribute('event','OnDocumentOpened(str,obj)')
      scriptStrs.textContent='OnDocumentOpened(str,obj)'
      this.ocx.appendChild(scriptStrs)
    },
    OpenDoc(){
      script.oframe.Open("C:\\Users\\cs2c\\Downloads\\L01_内核模块开发.docx", false, "Word.Document");
    }
  }
}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped>
h3 {
  margin: 40px 0 0;
}
ul {
  list-style-type: none;
  padding: 0;
}
li {
  display: inline-block;
  margin: 0 10px;
}
a {
  color: #42b983;
}
</style>
