/**
 * This plugin include wrapper of Bootstrap 5.2 components to be used in Vue JS (3) applications
 * It is designed to be loaded like a CDN plugin with no requirement for the code where it is 
 * installed to make any additional calls before mounting
 * 
 * There is an assumptions that a number of support libraries and style sheets will be loaded. These are:-
 * Vue (3.3.4)
 * VueUse
 * Bootstrap (5.2.3)
 * Ramda (0.29) 
 * Font Awesome (5.15)
 * 
 * The component included are:-
 * vbModal
 * vbOffcanvas
 * vbToast
 * RagSet
 * ConfirmAction
 * 
 * USAGE
 * 
 * vbModal
 * 
 * 
 * vbOffcanvas
 * 
 * 
 * vbToast
 *   <vb-toast ref="toast" message-title="Clock Event" message-body="It was stopped..." header-icon-class="fas fa-exclamation-circle text-danger"/>
 *   code to show it will be:
 *   this.$refs.<refName>.showToast(messageOverride, titleOverride, iconClassOverride)
 * 
 * RagSet 
 *  <rag-set :value="someTextPropertyWithRedAmberGreenValue" @change="method-to-set-value-red-amber-or-green"></rag-set>
 * 
 * ConfirmAction
 * <confirm-action ref="confirmWidget"
*    action-identifier="do-something"
 *   :collect-response-text="true" 
 *    message-detail="Why are you doing this?"
 *    message-title="Confirm some action"
 *    @button-click="confirmResult"
 * ></confirm-action>
 * and in some method....
 * someMethod () {
 *   this.$refs.confirmWidget.showConfirm()
 * }
 * confirmResult (actionName, buttonText, userText) {
 *   // code to handle here
 * }
 * 
 */



// During debugging & development, uncomment the following line to enable an intelligent editor to provide tooltips
// import Vue from "https://cdnjs.cloudflare.com/ajax/libs/vue/3.3.4/vue.global.min.js"
export default dhgBootstrapWrapperPack  = {
  install  (Vue, options) {
    
    Vue.component('vbToast', this.vbToast)
    Vue.component('vbModal', this.vbModal)
    Vue.component('vbOffcanvas', this.vbOffcanvas)
    Vue.component('sharePointFilePicker', this.sharePointFilePicker)
    Vue.component('RagSet', this.RagSet)
    Vue.component('ConfirmAction', this.ModalConfirmAction)
    Vue.component('GraphFolderButton', this.GraphFolderButton)
    Vue.component('LabelledControlGroup', this.LabelledControlGroup),
    Vue.component('vbCollapse', this.vbCollapse)
    Vue.component('vbTabSet', this.vbTabSet)
    Vue.component('vbPopover', this.vbPopover)
    
    Vue.component('WorkQueueButton', this.WorkQueueButton)
   
  },
  vbPopover: {
    props:{
      title: {type: String,  default: 'Popup Info'},
      message: {type: String,  default: 'Additional Information'},
      iconClasses: {type: String,  default: 'fas fa-info-circle'},
      buttonClasses: {type: String,  default: 'link-secondary '},
      tabindex: {type: Number,  default: 0},
      buttonText: {type: String,  default: ''}
    },
    computed: {
      iconClassesEvaluated() {
        let classes = this.iconClasses.replace(/\s+/g,'|').split('|')
        return classes
      },
      buttonClassesEvaluated() {
        let classes = this.buttonClasses.replace(/\s+/g,'|').split('|')
        return classes
      },

    },
    mounted() {
      const options = {
        trigger: 'focus'
      }
      const popover = new bootstrap.Popover(this.$refs.anchor, options) 
    },
    template: `
  <a ref="anchor" :tabindex="tabindex" 
    :class="buttonClassesEvaluated" role="button"
    data-bs-toggle="popover" data-bs-trigger="focus"
    :data-bs-title="title" 
    :data-bs-content="message" href="JAVASCRIPT:void(0)"><i :class="iconClassesEvaluated"></i>
  {{buttonText}}</a>
    `
  },

  vbTabSet: {
    inheritAttrs: false,
    data() {
      return {
        randomId: Math.random().toFixed(20).substring(2)
      }
    },
    props: {
      tabs: {type: Array, required: true, validator: (value, props)=>{
        if (value instanceof Array) {
          const fnNotNil = R.compose(R.not,R.isNil)
          if (R.any(R.isNil, value)){
            return false
          }
          return true
        } else {
          return false
        }
      }},
      modelValue: { required: true},
      removeDisabled: {type: Boolean, required: false, default: true},
      captionProperty: {type: String, required: false, default: 'caption'},
      codeProperty: {type: String, required: false, default: 'code'},
    },
    emits: ['tabSelected'],
    expose:['setTabCode'],
    methods:{
      setTabCode(code) {
        this.$emit('tabSelected', code)
      },
      shouldTabRender(tab) {
        const truthy = true
        if ((typeof tab === 'string' && tab.length) || this.removeDisabled !== true) return true;
        if (typeof tab === 'object' && tab.disabled == truthy ) return false;
        return true
      },
      isTabDisabled (tab) {
        const truthy = true
        if (typeof tab === 'string' && tab.length === 0)  return true;
        return  tab.disabled == truthy
      },
      
      tabCaption(tab){
        if (typeof tab === 'string') return tab;
        if (typeof tab === 'number') return tab;

        return tab ? tab[this.captionProperty]: ''

      },
      tabCode(tab){
        if (typeof tab === 'string') return tab;
        if (typeof tab === 'number') return tab;

        return tab ? tab[this.codeProperty]: ''
      },
      tabClass(tab){
        let aClasses = []
        let colour = this.tabCode(tab) === this.modelValue ? 'primary':'secondary'
        
        aClasses.push(`text-bg-${colour}`)
        
        if (this.isTabDisabled(tab)) {
          aClasses.push(`bg-${colour}-subtle`)
        }
        return aClasses      
      },
      tabInternalIdForIdx(tab, idx) {
        return `tab-${this.tabCode(tab)}-${idx}-${this.randomId}`
      }
    },
    mounted(){
      if(!(this.tabs instanceof Array)) throw new TypeError(`The tabs property must be specified as an array!`);
      let value = this.modelValue
      if (!value) { // auto select the first tab
        value = this.tabCode( this.tabs[0])
        this.setTabCode( value)
      }
    },
    template: `<ul class="nav nav-tabs">
          <transition-group name="simple-fade" >
          <template v-for="(tab,idx) in tabs" :key="tabInternalIdForIdx(tab, idx)" >
          <li v-if="shouldTabRender(tab)"
              class="nav-item pe-1"
              :key="tabInternalIdForIdx(tab, idx)"
              :aria-selected="tabCode(tab)===modelValue"
          >
            <a v-if="isTabDisabled(tab)"  class="nav-link"  v-html="tabCaption(tab)" :class="tabClass(tab)"></a>
            <a v-else class="nav-link" href="JAVASCRIPT:void(0)"
              @click="setTabCode(tabCode(tab))" v-html="tabCaption(tab)" :class="tabClass(tab)"></a>
          </li>
          </template>
          </transition-group>
        </ul>`

  },
  vbModal: {
    //name: 'V3BootstrapModal',
    inheritAttrs: false,
    data() {
      
      return {
        activeShown: 0, // Keeps track of how many calls,
        activeModal: null, // Will be populated after a SHOW,
        randomId: 'dialog' + Math.random(32).toString().substring(2)
      }
    },
    props: {
      title: {type: String, required: true, default: 'Modal Dialog'},
      titleIsHtml: {type:Boolean, required: false, default: false},
      scrollable: {type:Boolean, required: false, default: false},
      hideFooterButtons: {type:Boolean, required: false, default: false},
      size: {type: String, required: false, default: ''}, // can be "small", "large", "extra large", anything else gets default
      suppressDefaultClose: {type: Boolean, required: false, default: false},
      headerClasses: {type: String, required: false, default: ''},
      bodyClasses: {type: String, required: false, default: ''},
      footerClasses: {type: String, required: false, default: ''}
    },
    computed: {
      modalStyle: function ()  {
        let isSmall = /small/i.test(this.size)
        let isExtraLarge = /(extra).*(large)/i.test(this.size)
        let isLarge = /large/i.test(this.size)
        let isFull = /full/i.test(this.size)
        return  {
          'modal-dialog': true,
          'modal-dialog-scrollable': !!this.scrollable,
          'modal-sm': isSmall,
          'modal-lg': isLarge && !isExtraLarge,
          'modal-xl': isExtraLarge,
          'modal-fullscreen': isFull,

        }
      }
    },
    
    methods: {
      displayClassFromText (displayClassText) {
        const reTokens = /([a-z0-9-_]+ ?)/gi

        let out = []
        let mCol
        while(mCol = reTokens.exec(displayClassText)) {          
          out.push (mCol[1])
        }
        return out
      },
      showModal: function() {
        const options = {
          focus:true,
          keyboard: false,
          backdrop: 'static'
        }
        let modal = new bootstrap.Modal(this.$refs.thisModalDialog, options)
        this.activeModal = modal
        modal.show()
        this.$emit('open')

      },
      hideModal: function() {
        if (this.activeModal) {
          this.activeModal.hide()
          this.$emit('close')
          this.activeModal = null
        }
      },
      thisClose: function () {
        this.$emit('close')
      }
    },
    expose: ['showModal', 'hideModal'],
    emits: ['close','open'],
    template: `<div class="modal fade" tabindex="-1" ref="thisModalDialog"
      :aria-labelledby="randomId" aria-hidden="true"
      data-bs-backdrop="static" data-bs-keyboard="false">
      <div :class="modalStyle">
        <div class="modal-content">
          <div class="modal-header" :class="displayClassFromText(headerClasses)">
            <h5 class="modal-title" v-if="titleIsHtml" :html="title" :id="randomId">{{title}}</h5>
            <h5 class="modal-title" v-else :id="randomId">{{title}}</h5>
            <button v-if="!suppressDefaultClose" type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close" @click.prevent="thisClose"></button>
          </div>
          <div class="modal-body" :class="displayClassFromText(bodyClasses)" >
            <slot></slot>
          </div>
          <div class="modal-footer" v-if="!hideFooterButtons"  :class="displayClassFromText(footerClasses)" >
            <button type="button" class="btn btn-secondary" data-bs-dismiss="modal" @click.prevent="thisClose">Close</button>
          </div>
        </div>
      </div>
    </div>`

  },
  vbCollapse: {
    inheritAttrs: false,
    data() {
      
      return {
        active: false, 
        
        randomId: 'collapse-panel-' + Math.random(32).toString().substring(2)
      }
    },
    props: {
      // about the button used to open the canvas
      buttonCaption: {type: String, required: false, default: 'More ...'},
      showAtStart: {type: Boolean, required: false, default: false},
      buttonIconClass: {type: String, required: false, default: 'fas fa-angle-up'}, // Font Awesome Icon class text
      buttonIconClassActivated: {type: String, required: false, default: 'fas fa-angle-down'}, // Font Awesome Icon class text
      buttonClass: {type: String, required: false, default: 'btn btn-primary'}, // Bootstrap class text
      buttonStyle: {type: String, required: false, default: ''}, // for weird custom stuff
      buttonTitle: {type: String, required: false, default: 'More ...'} // Hover title
      
    },
    computed: {
      buttonActualClass() {
        let aClass = []
        const reTokens = /[\w\-_]+/gi
        let mCol
        while( mCol= reTokens.exec(this.buttonClass))
          aClass.push(mCol[0]);
        
        return aClass
      },
      buttonActualStyle() {
        let aStyles = []
        const reTokens = /[^;\w\-_]+/gi
        let mCol
        while( mCol= reTokens.exec(this.buttonStyle))
          aStyles.push(mCol[0]);
        return aStyles
      },
      buttonIconClosed() {
        let aClass = []
        const reTokens = /[\w\-_]+/gi
        let mCol
        while( mCol= reTokens.exec(this.buttonIconClass))
          aClass.push(mCol[0]);
        
        return aClass
      },
      buttonIconOpen() {
        let aClass = []
        const reTokens = /[\w\-_]+/gi
        let mCol
        while( mCol= reTokens.exec(this.buttonIconClassActivated))
          aClass.push(mCol[0]);
        
        return aClass
      }
    },
    methods: {
      toggle() {
        let elmPanel = document.getElementById(this.randomId)
        bootstrap.Collapse.getOrCreateInstance(elmPanel).toggle()
      }
    },
    mounted() {
      const divCollapsible = document.getElementById(this.randomId)
      divCollapsible.addEventListener('hidden.bs.collapse', event => {
        this.active=false
      })
      divCollapsible.addEventListener('shown.bs.collapse', event => {
        this.active=true
      })
      
    },
    template: `
    <button :data-bs-target="'#' + randomId"
      :class="buttonActualClass"
      :style="buttonActualStyle"
      :title="buttonTitle"
      @click="toggle"
    >{{buttonCaption}} 
       <span v-if="active"><i :class="buttonIconOpen"></i></span>
       <span v-else><i :class="buttonIconClosed"></i></span>
    </button>
    <div class="collapse" :id="randomId">
    <slot></slot></div>
    `
  },
  vbOffcanvas: {
    inheritAttrs: false,
    data() {
      
      return {
        activeShown: 0, // Keeps track of how many calls,
        activeCanvas: null, // Will be populated after a SHOW,
        randomId: 'offcanvas-' + Math.random(32).toString().substring(2)
      }
    },
    props: {
      // about the button used to open the canvas
      buttonCaption: {type: String, required: false, default: 'Show'},
      buttonIconClass: {type: String, required: false, default: ''}, // Font Awesome Icon class text
      buttonClass: {type: String, required: false, default: 'btn btn-primary'}, // Bootstrap class text
      buttonStyle: {type: String, required: false, default: ''}, // for weird custom stuff
      buttonTag: { validator(value, props) {return ['a','button'].includes(value) }, default:'button'},
 

      title: {type: String, required: false, default: 'Offcanvas Title'}, // Use when the header slot is missing
      // about the Offcanvas panel behaviour
      scroll: {type:Boolean, required: false, default: false}, // Allow the page behind to scroll when canvas is open (default no)
      keyboard: {type:Boolean, required: false, default: true}, // The ESC key closes the canvas when true (default)
      placement: {type:String, required: false, default: 'start'} // Which side the canvas appears into view
    },
    expose: ['show', 'hide'],
    emits: ['close', 'click', 'open', 'showOffcanvas', 'hideOffcanvas'], // The click is for the button that controls the canvas
    computed: {
      canvasClass: function () {
        let placement
        switch (this.placement) {
          case 'start':
          case 'end':
          case 'top':
          case 'bottom':
            placement = this.placement
            break;
          case 'right':
            placement = 'end'
          default:
            placement = 'start'

        }
        return {
          'offcanvas-start': placement === 'start',
          'offcanvas-end': placement === 'end',
          'offcanvas-top': placement === 'top',
          'offcanvas-bottom': placement === 'bottom',
          offcanvas: true
        }
      }
    
    },
    methods: {
      mainButtonClick: function (evt) {
        // Open the panel after raising a click event (in case the parent needs to do stuff in the slot!
        this.activeShown ++
        this.$emit('click', evt)
        this.show()
      },
      show: function() {
        if (!this.activeCanvas) {
          const options = {
            keyboard: this.keyboard,
            scroll: this.scroll
          }
          let domElement = this.$refs.myCanvas
          let canvas = new bootstrap.Offcanvas(domElement, options)
          this.activeCanvas = canvas
        }
        this.activeCanvas.show()
        this.$emit('showOffcanvas')

      },
      hide: function() {
        if (this.activeCanvas) {
          this.activeCanvas.hide()
          this.$emit('hideOffcanvas')
        }
      }
    },
    errorCaptured: (err, vm, info) => {
        
      console.warn(`Error raised i the Vue application chain...\n${info}`)
      console.log(vm)
      console.log(err)
      return false // stop bubbling further
    },

    template: `<div  tabindex="-1" ref="thisCanvasSet" :class="{'d-inline-block': buttonTag==='a'}">
    
      <button v-if="buttonTag !== 'a'" role="button" class="btn" :class="buttonClass" :style="buttonStyle" @click="mainButtonClick($event)">
        <i v-if="buttonIconClass" :class="buttonIconClass"></i> {{buttonCaption}}
      </button>
      <a v-else :class="buttonClass" :style="buttonStyle" @click="mainButtonClick($event)">
        <i v-if="buttonIconClass" :class="buttonIconClass"></i> {{buttonCaption}}</a>
      <Teleport to="body">
        <div ref="myCanvas" :class="canvasClass" tabindex="-1" id="offcanvasRight" aria-labelledby="offcanvasRightLabel">
          <div class="offcanvas-header">
            <slot name="header">{{title}}</slot>
            <button type="button" class="btn-close" aria-label="Close" @click="hide"></button>
          </div>
          <div class="offcanvas-body"><slot>...</slot></div>
          <div class="offcanvas-footer"><slot name="footer"></slot></div>
        </div>
      </Teleport>
    </div>`

  },
  vbToast: {
    inheritAttrs: false,
    data() {
      
      return {
        activeShown: 0, // Keeps track of how many calls,
        fnCancelTimer: null,
        activeToast: null, // Will be populated after a SHOW

        defaultMessage: '', // Take from Props and use when not overridden in the show method
        defaultTitle: '',
        defaultIcon: '',
        
        activeMessage: '', // Take from show method ars
        activeTitle: '',
        activeIcon: '',
        randomId: 'toast-' + Math.random(32).toString().substring(2)

      }
    },
    props: {
      messageTitle: {type: String, required: true, default: 'I toast you ...'},
      messageBody: {type: String, required: false, default: '<span></span>'},
      headerIconClass: {type: String, required: false, default: 'fas fa-exclamation-circle text-warning'}, // class for an <i></i> to display an icon
      positionalLocationClass: {type: String, required: false, default: 'position-absolute top-0 start-5'}, 
      duration: {type:Number, required: false, default: 3000} // HOw many milliseconds to show before dismissal
    },
    created: function () {
      this.defaultMessage = this.messageBody
      this.defaultIcon = this.headerIconClass,
      this.defaultTitle = this.messageTitle
    },
    methods: {
      dismissAuto: function() { // Invoked either by a timeout (so the timer is finished) 
        let toastBootstrap = this.activeToast
        if (toastBootstrap && toastBootstrap.isShown()) {
          toastBootstrap.hide()
          this.activeToast = null
        }
        this.activeShown --
        this.activeToast =  null
      },
      dismissToast: function ($event) {
        if (typeof this.fnCancelTimer === 'function') {
          this.fnCancelTimer()
          this.fnCancelTimer = null
        }
        this.dismissAuto()
      },
      clearTimer: function(hTimeout) {
        window.clearTimeout(hTimeout)
      },
      showToast: function(message = '', title = '', iconClass = '') {


        // Clean up any prior timeouts
        if (this.activeShown) {
          // Already been displayed and NOT yet hidden so just kill any cancel timer
          if (typeof this.fnCancelTimer === 'function') {
            this.fnCancelTimer()
          }
        }

        this.activeMessage = message ? message : this.defaultMessage
        this.activeTitle = title ? title : this.defaultTitle
        this.activeIcon = iconClass ? iconClass : this.defaultIcon
        // Set up the Toast ONLY if the REF is ready ....
        if (this.$refs.elToast !== null) {
          let toastBootstrap = new bootstrap.Toast(this.$refs.elToast)
          toastBootstrap.show()
          let msToShow = this.duration
          if (isNaN(msToShow) || msToShow <  100 ){
            msToShow = 3000 // default  3 seconds if try less 1/20th
          }

          toastBootstrap.show()
          this.activeShown ++
          this.activeToast = toastBootstrap

          // Prepare to Auto CLose  using a timer
          let hTimeout = window.setTimeout(()=>{}, 50) // Do nothing, but lets us know what the  NEXT timeout handle is
          
          this.fnCancelTimer = this.clearTimer.bind(this, hTimeout+1)
          let fnCompositeCleanup = R.compose(
            this.dismissAuto.bind(this),
            this.fnCancelTimer
          )// A composition of 2 functions to cancel the time
          window.setTimeout(fnCompositeCleanup, msToShow )
          
          
        }
      },
    },
    expose: ['showToast'],
    template: `<div class="toast-container opacity-100" :class="positionalLocationClass">
      <div ref="elToast" class="toast" role="alert" aria-live="assertive" aria-atomic="true">
        <div class="toast-header bg-secondary  text-white opacity-100">
          <span class="pe-1"><i :class="activeIcon"></i></span>
          <strong class="me-auto">{{ activeTitle }}</strong>
          
          <button type="button" class="btn-close" @click.prevent="dismissToast" aria-label="Close"></button>
        </div>
        <div class="toast-body opacity-100 bg-light" v-html="activeMessage">       
        </div>
      </div>
    </div>`
  },
  sharePointFilePicker: {
    name: 'sharePointFilePicker',
    inheritAttrs: true,
    data() {
      
      return {
        activeShown: 0, // Keeps track of how many calls,
        popupWindow: null, // Will be populated after a SHOW,
        randomId: 'dialog' + Math.random(32).toString().substring(2),
        fnMessageListener: null
      }
    },
    props: {
      listenOnly: {type: Boolean, required: false, default: true},
      siteUrl: {type: String, required: true },
      folderPath: {type: String, required: false, default: 'Shared Documents/'},
      title: {type: String, required: false, default: 'Select a file ..'},
      buttonCaption: {type: String, required: false, default: 'File ...,'},
      buttonIcon: {type: String, required: false, default: 'fa-folder-open'},
      buttonClass: {type: String, required: false, default: 'btn-primary'},
      // startFolder: {type:String, required: false, default: '/Shared Documents'},
      tagButtonOrAnchor: {type:String, required: false, default: 'button'},
      typeFilters: {type:String, required: false, default: ''}
    },
    computed: {

      pickerUrl: function () {
        
        const siteBase = this.siteUrl
        const folderPath = this.folderPath
        const reGetSiteRelPath = /^https:\/\/[^:\/]+(:[0-9]+)?\/(.*?)(\/[0-9]+)?(\?.*)?$/;
        // const serverRelativePath = siteBase.startsWith('https://') ? siteBase.replace(reGetSiteRelPath, '$2') : siteBase
        let encodedId = `${encodeURI(siteBase)}/${encodeURI(folderPath)}`
        let urlTypeFilters = this.typeFilters.trim() ? `&typeFilter=${this.typeFilters}` :''
        return `${siteBase}/_layouts/15/OneDrive.aspx?p=2&id=${encodedId}${urlTypeFilters}`
      },
      derivedButtonClass: function () {
        let oClass  = { btn : true }
        let aClasses = this.buttonClass.replace(/\s+/g,'#').split('#')
        aClasses.forEach(sClass => {
          oClass[sClass] = true
        })

        return oClass
      }
    },
    methods: {
      receiveMessage: function (evt) {
        if (!evt.isTrusted) return ;
        const PICKER_EVENT = '[OneDrive-FromPicker]'
        
        try {
          let eventData = evt.data

          if (('' + eventData).startsWith(PICKER_EVENT)) {
            
            let pickedData = JSON.parse(eventData.substring(PICKER_EVENT.length))
            switch (pickedData.type) {
              case 'cancel':
                this.$emit('cancelFilePick')
                this.cancelPopup()
                break
              case 'success':
                console.log('one drive picker SUCCESSm 1st file will be used ...')
                console.log(pickedData.items)

                this.$emit('filePicked', pickedData.items[0])
                this.cancelPopup()
                break
              default:
                break // Ignore the many messages produced by SharePoint that are not interesting!
            }
            
          }
        } catch (ex) {
          console.error(`Unexpected error in message receive handler ...\n${ex.message}`,evt)
          
        }
      },
      showPicker: function() {
        try {
          this.startListening() // For messages about file  selections
          if (this.listenOnly === false) {
            this.showPickerWindow()
          }
          this.$emit('listenToUrl', this.pickerUrl)
          
        } catch (ex) {
          console.error(`Failure from chain of SharepointFilePicker::showPicker\nERROR: ${ex.message}`, ex)
        }
      },
      showPickerWindow: function (windowFeatures = 'popup') {
        let winPicker = window.open(this.pickerUrl, 'SharePoint-Picker', windowFeatures)
        winPicker.focus()
        this.popupWindow = winPicker
      },
      startListening: function () {
        let fnMessageListener = this.receiveMessage.bind(this)
        this.fnMessageListener = fnMessageListener
        console.log('Listening for messages post from OneDrive.aspx page')
        window.addEventListener('message', fnMessageListener, false)
      },
      stopListening: function () {
        if (this.fnMessageListener) {
          console.log('Stop listening for messages post from OneDrive.aspx page')
          try {
            window.removeEventListener('message', this.fnMessageListener)
            this.fnMessageListener = null
          } catch (ex) {
            console.error(`Failed to remove the message listener`)
          }
        }

      },
      cancelPopup: function() {
        
        this.stopListening()
        if (this.listenOnly === false) {
          this.cancelPopupWindow()
        }
      },
    
      cancelPopupWindow: function () {
        if (this.popupWindow) {
          try{
            this.popupWindow.close()
          } catch (ex) {
            console.error
          }
          this.popupWindow = null
          

        } else {
          console.warn('Popup Picker is not active!')
        }
      }
    },
    emits: ['filePicked', 'cancelFilePick', 'listenToUrl'],
    expose: ['showPicker', 'cancelPopup', 'startListening'],
    template: `
  
      <button  v-if="tagButtonOrAnchor === 'button'" class="btn" :class="derivedButtonClass" @click.prevent="showPicker">
        <i v-if="buttonIcon" :class="['fas', buttonIcon]"></i>
        {{buttonCaption}}
      </button>
      <a v-else role="button" href="JAVASCRIPT:void(0)"  @click.prevent="showPicker">
        <i v-if="buttonIcon" :class="['fas', buttonIcon]"></i> {{buttonCaption}}
      </a> 
  
  `
  },
  RagSet: {
    data() {
      return ({
        internalStatus: '',
        isRed: false,
        isAmber: false,
        isGreen: false
      })
    },
    
    props: {
      ragCaptions: {type: Array, required: false, default: ['Red','Amber', 'Green'] },
      modelValue: { required: true},
     
      disabled: {type:Boolean, required: false, default: false},
      offOpacity: {type: Number, required: false, default: 50},
      title: {type: String, required: false, default:  'Generic RAG'},
      allowEmptyToggle: {type: Boolean, required: false, default: true},
      mWidth: {type: Number, required: false, default: 5}
    },
    watch:{
      modelValue: {
        handler(newValue, oldValue) {
          this.setStatesFromValue(newValue)
        },
        immediate: true
      }
    },
    emits: ['change','update:modelValue'],
    computed:{
      redClass () { return this.getButtonClasses('danger', this.isRed) },
      amberClass () { return this.getButtonClasses('warning', this.isAmber) },
      greenClass () { return this.getButtonClasses('success', this.isGreen) },
      buttonStyle () {
        return { width: this.mWidth + 'em' }
      }
    },
    methods:{
      setStatesFromValue(value) {
        let idxButton = this.ragCaptions.findIndex(R.equals(value))
        this.isRed = idxButton === 0
        this.isAmber = idxButton === 1
        this.isGreen = idxButton === 2
      },
      getButtonClasses(bootstrapColour, isActive) {
        let aClass=['btn', 'font-monospace','btn-sm',  'btn-' + bootstrapColour]
        if (isActive) {
          aClass.push('text-decoration-underline', 'font-bold', 'shadow')
        } else {
          
          let allowedOpacity = [0,25,50,75,100]
          let offOpacity = allowedOpacity.find(R.equals(this.offOpacity)) 
          aClass.push('opacity-' +  offOpacity.toString())
        }
        return aClass
      },
      raiseChange (btnIndex) {
        if (this.disabled) return;
        let changeTo = this.ragCaptions[btnIndex]
        if (this.internalStatus === changeTo) {
          // Clicked on button for the CURRENT state
          if (this.allowEmptyToggle) {
            changeTo = null
            this.$emit('change', changeTo)
            this.$emit('update:modelValue', changeTo)
          }
        } else {
          this.$emit('change', changeTo)
          this.$emit('update:modelValue', changeTo)
        }
        this.internalStatus = changeTo
      }
    },
    
    mounted (){
      this.internalStatus =  this.modelValue
    },
    template: `
<div class="btn-group"  role="group" :aria-label="title + ' RAG Button Group'">
  <button type="button" 
  :style="buttonStyle"
  :class="redClass"   
  :title="'Change ' + title + ' to ' + ragCaptions[0]"
  :aria-label="'Button set ' + title + ' to ' + ragCaptions[0]"
  @click="raiseChange(0)" 
  >
    {{ragCaptions[0]}}
  </button>
  <button type="button" 
      :style="buttonStyle"
      :class="amberClass"
      :title="'Change ' + title + ' to ' + ragCaptions[1]"
      :aria-label="'Button set ' + title + ' to ' + ragCaptions[1]"
      @click="raiseChange(1)"
    >
      {{ragCaptions[1]}}
    </button>
    <button type="button" 
      :style="buttonStyle"
      :class="greenClass"
      :title="'Change ' + title + ' to ' + ragCaptions[2]"
      :aria-label="'Button set ' + title + ' to ' + ragCaptions[3]"
      @click="raiseChange(2)" 
    > 
    {{ragCaptions[2]}}
  </button>
</div>    
    `
  },

  ModalConfirmAction: {
    props: {
      actionIdentifier: {type: String, required: true},
      buttons: {type:Array, required: false, default: ['OK', 'Cancel']},
      buttonClasses: {type:Array, required: false, default: ['btn btn-primary', 'btn btn-secondary']},
      messageTitle: {type: String, required: false, default: 'Confirm'},
      messageDetail: {type: String, required: false, default: ''},
      collectResponseText: {type: Boolean, required: false, default: false},
      collectResponseLabel: {type: String, required: false, default: 'Your Response'}
    },
    expose: [ 'showConfirm'],
    emits:['buttonClick'],
    data() {
      
      return { // Almost No own data
        userResponseText: '',
        currentActionIdentifier: '',
        currentTitle: '',
        currentDetail: '',
        currentUserTextLabel: '',
        randomId: 'dialog' + Math.random(32).toString().substring(2)
      }
    },
    watch:{
      messageTitle(value) {this.currentTitle = value},
      messageDetail(value) {this.currentDetail = value},
      collectResponseLabel(value) {this.currentUserTextLabel = value},
      actionIdentifier(value) { this.currentActionIdentifier = value}
    },
    mounted () {
      this.currentDetail = this.messageDetail,
      this.currentTitle = this.messageTitle,
      this.currentUserTextLabel = this.collectResponseLabel
      this.currentActionIdentifier = this.actionIdentifier
    },
    methods: {
      showConfirm(titleOverride = '', detailOverride = '', userTextLabel = '') {
        if (titleOverride) this.currentTitle = titleOverride;
        if (detailOverride) this.currentDetail = detailOverride;
        if (userTextLabel) {
          this.currentUserTextLabel = userTextLabel
        }
        this.userResponseText = ''
        this.$refs?.dlgConfirm?.showModal()
      },
      clickUserButton(btnText) {
        this.$emit('buttonClick', this.actionIdentifier, btnText, this.userResponseText)
        
        this.$refs?.dlgConfirm?.hideModal()

        // Restore defaults ...
        this.currentDetail = this.messageTitle
        this.currentTitle = this.messageTitle
        this.currentUserTextLabel = this.collectResponseLabel

        
      },
      getButtonClass(idx) {
        let textClasses = this.buttonClasses[idx]
        const reTokens = /([a-z0-9-_]+ ?)/gi

        let out = ['me-2']
        let mCol
        while(mCol = reTokens.exec(textClasses)) {
          
          out.push (mCol[1])
        }
        return out

      }
    },
    template: `
    <vb-modal 
      header-classes="bg-dark-subtle text-emphasis-dark" body-classes="bg-dark-subtle text-emphasis-dark"
      ref="dlgConfirm"
      :title="currentTitle"
      :hide-footer-buttons="true"
      :suppress-default-close="true">
      <div v-if="currentDetail" class="lead" style="white-space: pre-wrap">{{currentDetail}}</div>
      <div v-if="collectResponseText" class="form-floating mb-2">
        <textarea class="form-control w-100" v-model="userResponseText" style="height: 4rem" cols="40" rows="6" :id="'userText'+ randomId"></textarea>
        <label :for="'userText'+ randomId">{{currentUserTextLabel}}</label>
      </div>
      <div class="float-end">
        <button v-for="(btn, idx) in buttons" :key="'btn:' + btn" @click="clickUserButton(btn)" 
          :class="getButtonClass(idx)"
          role="button"
        >{{btn}}</button>

      </div>
    </vb-modal>
    `
  },
  WorkQueueButton: {
    props:{
      siteUrl: {type:String, required: false, default: undefined},
      queueList: {type:String, required: false}, // Name or Id
      statusField: {type:String, required: false, default:'status'},
      optionsField: {type:String, required: false, default:'options'},
      processLogField: {type:String, required: false, default:'processLog'},
      successStatus: {type:String, required: false, default:'SUCCESS'},
      failedStatus: {type:String, required: false, default:'FAILED'},
      readyToSend: {type: Boolean, required: false},
      buttonText: {type: String, required: false, default: 'Send Request'},
      requestTitle: {type: String, required: false, default: undefined},
      requestOptions: {type: Object, required: true},
      attachmentFiles: {type: Array, required: false, default: []}
    },
    components:{
      
    },
    expose: ['reset'],
    emits: ['request-sent','request-success', 'request-failed'],
    data() {
      return { // No real own data - all derived from properties
        queueItemId: null, // Number
        isPolling: false,
        pollingHandle: null,
        requestSent: false,
        requestComplete: false,
        requestFailed: false,
        requestorName: null,
        queueStatus: 'Not Sent',
        queueProcessLog: null,
        targetSiteUrl: null
        
      }
    },
    watch: {
      queueProcessLog(newText) {
        
        this.$refs.logDiv.innerHTML = newText || '-- No Log --'
      }
    },
    computed: {
      requestDateText() {
        const fmtDate = new Intl.DateTimeFormat('en-GB')
      },
      isDisabled() { return this.readyToSend !== true || this.requestSent },
      buttonTooltip(){
        return `${this.buttonText} for ${this.requestTitle}`
      },
      queueItemTitle () {
        return this.requestTitle || `Request by ${this.requestorName} on ${moment().format('Do MMM YY at h:mm A')}`
      },

      mapFields() {
        // Map Fields is a 2D array suited for use by the maceSPListUtility library.
        // It is normally used as an abstraction layer to between a logical "module" amd physical list layout of SharePoint
        // In this case the specific columns that might be needed are derived from properties about the list
        // We will always need to "set" and "read" some core columns whose actual implementation details can vary .
        // The grid array here keeps the interface we deal with clean (and abstracted)
        return [
          ['Title', 'requestTitle', 'Request Title', 'Text'],
          [this.statusField || 'status','status','Status','Text'],
          [this.optionsField || 'options', 'options', 'Options (JSON)', 'Note'],
          [this.processLogField || 'processLog', 'processLog', 'Note']
        ]

      },
      statusBadgeClass() {
        let aClass=[]
        if (this.queueStatus === 'WAITING') {
          aClass.push('text-bg-secondary')
        } else if(this.requestComplete && !this.requestFailed) {
          aClass.push('text-bg-success') 
        } else if (this.requestFailed) {
         aClass.push('text-bg-danger') 
        } else {
          aClass.push('text-bg-info') 
        }
        return aClass
      }

    },
    methods: {
      async pollQueue() {
        console.assert(this.queueItemId > 0, `Polling for queue status aborted because no item # available`)
        if (this.queueItemId) {
          let oQueueEntry = (await maceSPListUtility.getSiteListItems(this.queueList,this.mapFields,{itemLimit: 1, filterClause: 'Id eq ' + this.queueItemId}, this.targetSiteUrl))[0]
          let oConvertedItem = maceSPListUtility.convertSharePointListToLogicalFromFieldMap(this.mapFields, oQueueEntry)
          this.queueStatus = oConvertedItem.status
          this.queueProcessLog = oConvertedItem.processLog
          if (oConvertedItem.status === this.successStatus) {
            this.$emit('request-success')
            this.requestComplete = true
            this.stopPolling()
          }
          if (oConvertedItem.status === this.failedStatus) {
            this.$emit('request-failed')
            this.requestComplete = true
            this.requestFailed = true
            this.stopPolling()
          }

        }
      },
      
      startPolling() {
        this.isPolling = true
        this.pollingHandle =  window.setInterval(this.pollQueue.bind(this), 3000)
      },
      stopPolling() {
        if (this.isPolling) {
          window.clearInterval(this.pollingHandle)
          this.isPolling = false
        }        
      },
      reset () {
        this.stopPolling()
        this.requestSent = false
        this.queueItemId = null
        this.queueStatus =  'Not Sent'
        this.requestComplete = false
        this.queueProcessLog = null
        this.requestFailed = false
      },
      async sendRequest () {
        
        let payload = {
          requestTitle: this.queueItemTitle,
          status: 'WAITING',
          options: JSON.stringify(this.requestOptions),
          processLog: `<h3>Processing Requested</h3><p>For: ${this.queueItemTitle}</p>`
        }
        let oResponse = await maceSPListUtility.createListItemUsingPhysicalToLogicalMapFullResponse(this.queueList, this.mapFields, payload, this.targetSiteUrl)
        this.queueItemId = oResponse.Id
        let oConvertedItem = maceSPListUtility.convertSharePointListToLogicalFromFieldMap(this.mapFields, oResponse)
        this.queueStatus = oConvertedItem.status
        this.queueProcessLog = oConvertedItem.processLog
        this.requestSent = true
        this.startPolling()
        if (this.attachmentFiles && this.attachmentFiles instanceof Array && this.attachmentFiles.length) {
          for (let i = 0; i < this.attachmentFiles.length; i++){
            let file = this.attachmentFiles[i]
            try {
              await maceSPListUtility.addAttachmentToListItem(this.queueList, this.queueItemId, file, this.targetSiteUrl)
            } catch (error) {
              console.error(`Failed to process attachment #${i + 1}, error: ${error instanceof Error ? error.message : error?.toString()}`)
            }
          }
        }

      },
      showLog() {
        this.$refs.dlgLog.showModal()
      }
    },
    created() {
      maceSPListUtility.getCurrentUserInSite().then(u=> this.requestorName = u.Title)
      if (this.siteUrl) {
        this.targetSiteUrl = this.siteUrl

      } else {
        maceSPListUtility.getContextInfoForSite().then(d=> this.targetSiteUrl = d.WebFullUrl)
      }
    },
    mounted () {
      
    },
    template: `<div class="d-inline-block  position-relative" >
    <div class="btn-group" role="group" aria-label="Request Buttons">
      <button type="button" @click="sendRequest" class="btn btn-sm btn-primary" :disabled="isDisabled">
        <span v-if="!requestSent"><i class="fas fa-plus-circle"></i></span>
        <span v-if="requestSent && !requestComplete"><i class="fas fa-spinner fa-spin"></i></span>
        <span v-if="requestComplete && !requestFailed"><i class="fas fa-check-circle"></i></span>
        <span v-if="requestFailed"><i class="fas fa-times-circle text-bg-danger"></i></span>
         {{buttonText}}
      </button>
      <button type="button" @click="showLog" class="btn btn-sm btn-secondary" :disabled="!requestSent"><i class="fas fa-clipboard"></i> Log</button>
      <span v-if="requestSent" :class="statusBadgeClass" class="badge  position-absolute top-0 start-0 translate-middle">{{queueStatus}}</span>
    </div>
    <vb-modal title="Queue Process Log" ref="dlgLog" size="large" :scrollable="true">
    <div ref="logDiv">...</div>
    </vb-modal>
    
  </div>`
  },

  GraphFolderButton: {
    props:{
      siteUrl: {type:String, required: true},
      driveId: {type:String, required: false},
      folderId: {type:String, required: false},
      target: {type:String, required: false, default: '_blank'},

    },
    data() {
      return { // No real own data - all derived from properties
        webUrl: '',
        folderName: '-',
        childCount: null
      }
    },
    watch: {
      graphUrl (newValue) {
        this.evaluateUrl(newValue)
      }
    },
    computed: {
      graphUrl() {
        if (this.siteUrl && this.driveId && this.folderId) {
        return `${this.siteUrl}/_api/v2.1/drives/${this.driveId}/items/${this.folderId}`
        } else {
          return ''
        }
      },
      isDisabled() { return this.webUrl.length===0}
    },
    methods: {
      async evaluateUrl(url) {
        try {
          const resp = await fetch(url, {headers:{Accept: 'application/json'}})
          const payload = await  resp.json()
          if (payload ) {
            if (payload.hasOwnProperty('folder')) {
              this.webUrl = payload.webUrl
              this.folderName = payload.name
              this.childCount = payload.folder.childCount
            } else {
              this.webUrl = ''
              this.folderName = '-'
              this.childCount = null
            }
              
          }
        } catch (error) {
          console.error('Failed attempting to validate Graph API detail.')
          console.error(error)
        }
      },
      openFolder () {
        window.open(this.webUrl, this.target)
      }
    },
    mounted () {
      if (this.graphUrl) this.evaluateUrl(this.graphUrl);
    },
    template: `<div class="d-inline-block text-truncate position-relative" >
    <button  @click="openFolder" class="btn btn-link" :disabled="isDisabled">
      <span v-if="isDisabled"><i class="fas fa-ban fa-stack-1x"  style="color:Tomato"></i>- No Folder -</span>
      <span v-else><i class="fas fa-folder-open "></i> <span class="fs-6 text-muted">/{{folderName}}</span></span>
     
    </button>
    
  </div>`
  },

  LabelledControlGroup: {
    name: 'LabelledControlGroup',
    inheritAttrs: false,
    props: {
      label: {type: String, required: true},
      placeholder:{type: String, required: false},
      labelClass: {type: String, required: false, default: 'fs-5'}, // Label is fs-5 by default
      type: {
        validator(value, props) {
          return ['text','textarea','number', 'percent','date','color','datetime-local','currency','select', 'checkboxes', 'switch', 'radio', 'lookup', 'lookupMulti'].includes(value)
        }, default:'text'
      },
      required: {type: Boolean, required: false, default: false},
      requiredIconName: {type: String, required: false, default: 'fa-asterisk'}, // Optional display icon override
      modelValue: {required: true},
  
      options: {type: Array, required: false},
      optionBackgroundColourProperty: {type: String, required: false}, // Property name of options item to hold a hex colour
      optionForegroundColourProperty: {type: String, required: false}, // Property name of options item to hold a hex colour
      optionValueProperty: {type: String, required: false, default: ''}, // Property name to bind value to
      optionValuePropertyList: {type: Array, required: false, default: ['Id','ID','id', 'key','code', 'name']}, // Property names that *might* be an ID of an object (when not specified)
      optionTextProperty: {type: String, required: false, default: ''}, // Property name to use for display text
      optionFormatterFunction: {type: Function, required: false, default: null},
      optionBindObject: {type: Boolean, required: false, default: false},
      selectMulti: {type: Boolean, required: false, default: false},
      
      switchOnValue: {type: undefined, required: false, default: true},
      switchOffValue: {type: undefined, required: false, default: false},
      
      missingColourClass: {type: String, required: false, default: 'text-danger'},
      presentColourClass: {type: String, required: false, default: 'text-success'},
      readonly: {type: Boolean, required: false, default: false},
      suppressPrefixIcon: {type: Boolean, required: false, default: false},
      
      rows: {type: String, required: false, default: '3'},
      min: {type: Number, required: false},
      max: {type: Number, required: false},
      step: {type: Number, required: false},
     // fixedDecimals: {type: Number, required: false},
      maxlength: {type: String, required: false, default: '255'}, // Particularly for single line input
      currencyLocale: {type: String, required: false, default: 'en-GB'},
      currency: {type: String, required: false, default: 'GBP'},
  
      // For SharePoint lookup lists source
      lookupListName: {type: String, required: false},
      lookupListSite: {type: String, required: false},
    },
    emits: ['update:modelValue','checked'], // push out the standard edit event to the parent
    expose:['requiredPass'],
    data() {
      return {
        localId: null,
        selectedItemToAddGroup: null,
        boundSelectIndex: undefined,
        pushedBoundSelectedValue: false,
      }
    },
    watch:{
      boundSelectIndex:{
        handler(newIndex,oldIndex) {
          if (newIndex === oldIndex){ 
            console.log('unchanged boundSelectIndex triggered (watching abort!)')
            return 
          }
          if (this.pushedBoundSelectedValue) return 
          let newOption = this.options[newIndex]
          let emitValue = this.optionBindObject ? newOption : this.getBestOptionValue(newOption)
          this.$emit('update:modelValue', emitValue)
        }
      },
      modelValue:{
        immediate: true,
        handler(newValue, oldValue) {
          if (this.type === 'select') {
            let idx = this.getOptionIndex(newValue)
            this.pushedBoundSelectedValue = true
            
            if (idx > - 1){
              this.boundSelectIndex = idx
            } else {
              this.boundSelectIndex = undefined
            }
            this.$nextTick(()=> this.pushedBoundSelectedValue = false)
          }
        }
      }, 
      options: {
        flush:'post',
        handler(newSet, oldSet){
          // react to changes in options that may affect the selected index
          let idxNow = this.boundSelectIndex
          let idxAfter = this.getOptionIndex(this.modelValue)
          if (idxAfter !== idxNow && typeof idxAfter === 'number' && idxAfter> -1){
            this.pushedBoundSelectedValue = true
            this.boundSelectIndex = idxAfter
            this.$nextTick(()=> this.pushedBoundSelectedValue = false)
          }
        }
      }
    },
    
    computed:{
      
      haveValue() {
        let v = this.modelValue
        if (typeof v === 'string' &&  v === '') return false;
        return !(R.isNil(v))
      },
      typeIconClass() {
        let aClass= ['fas']
        let iconName = ''
        switch(this.type) {
          case 'text':
            iconName = 'fa-font'
            break;
          case 'textarea':
            iconName = 'fa-pen-fancy'
            break;
          case 'number':
            iconName = 'fa-hashtag'
            break;
          case 'percent':
            iconName = 'fa-percent'
            break;
          case 'date':
            iconName = 'fa-calendar'
            break;
          case 'color':
            iconName = 'fa-palette'
            break;
          case 'datetime-local':
            iconName = 'fa-clock'
            break;
          case 'currency':
            iconName = 'fa-pound-sign'
            break;
          case 'select':
            iconName = 'fa-list-ul'
            break;
          case 'checkboxes':
            iconName = 'fa-check-double'
            break
          case 'radio': 
            iconName = 'fa-check-circle'
            break
          case 'switch': 
            // Does not use an icon because display layout differs to standard types
            break
          default:
            break;
        }
        aClass.push(iconName)
        return aClass
      },
      invalidIconClass () {
        return  [
          'fas',
          this.requiredIconName,
          this.missingColourClass
        ]
      },
      validIconClass () {
        return  [
          'fas',
          this.requiredIconName,
          this.presentColourClass
        ]
      },
      effectiveLabelClass() {
        let textClasses = this.labelClass
        const reTokens = /([a-z0-9-_]+ ?)/gi
  
        let out = ['form-label']
        let mCol
        while(mCol = reTokens.exec(textClasses)) {
          out.push (mCol[1])
        }
        return out
      },
      
      requiredPass () {
        return this.required === false || this.haveValue
      },
      isDateType() {
        return this.type === 'date' || this.type === 'datetime-local'
      },
      isBooleanType () {
        let aTypes=['switch']
        return aTypes.includes(this.type)
      },
      isNumberType() {
        const aNumberTypes = ['number','percent','currency']
        return aNumberTypes.includes(this.type)
      },
      isCurrencyType() {
        const aNumberTypes = ['currency']
        return aNumberTypes.includes(this.type)
      },
      renderedControlId () {
        return `labelledControl-${this.localId}`
      },
      useGroup() {
        let aNonGroupLayoutTypes = ['radio','checkboxes','switch']
        let generateGroup = !aNonGroupLayoutTypes.includes(this.type)
        if (this.type ==='select' && this.selectMulti) generateGroup = false //will need alternate UI of a box containing "badges" of items
        return generateGroup
        
      },
      getActiveOptionStyle () {
        let oNoStyle ={}
        if (this.options?.length > 0 &&   this.optionBackgroundColourProperty || this.optionForegroundColourProperty) {
          let currentValue = this.modelValue
          let aOpts = this.options.filter(opt => this.getBestOptionValue(opt) === currentValue)
          if (aOpts.length) {
            return this.getOptionStyle(aOpts[0])
          }
        }
        return oNoStyle
      },
      modelValueAsNumber: {
        get(){
          if ( isNaN(this.modelValue)) {
            return undefined
          } else {
           let value = parseFloat(this.modelValue)
           return value //R.isNil(this.fixedDecimals) ? value :  value.toFixed(this.fixedDecimals)
          }
        },
        set(newValue) {
          let fValueRaw = null
          if (typeof newValue === 'string') {
            let reNonNumber = /[^\d\.\+\-]/g
            let strippedValue = ('' + newValue).replace(reNonNumber,'')
            fValueRaw = parseFloat(strippedValue)
          } else {
            fValueRaw = typeof newValue === 'number' ? newValue : parseFloat(newValue)
          }
          
          if (isNaN(fValueRaw)){
            this.$emit('update:modelValue', null)
            this.$refs.num.value = ''
            
          } else {
            let min = R.isNil(this.min) ? Number.NEGATIVE_INFINITY : this.min
            let max = R.isNil(this.max) ? Number.POSITIVE_INFINITY : this.max
            let fValueClamped = R.clamp(min, max, fValueRaw)
          
            
            if (this.$refs.num.value != fValueClamped ) {
              this.$refs.num.value = fValueClamped
            }
            this.$emit('update:modelValue', fValueClamped)
          }
  
        }
      },// () {return  parseFloat(this.modelValue)},
      /**
       * @description display the model value multiplied by 100 to express a decimal as a percentage
       * This works un conjunction with the update logic that divides by 100
       */
      modelValueAsPercentage: {
        get(){
          let fValueRaw = this.modelValueAsNumber
          if (isNaN(fValueRaw)) return  ''
          //let min = R.isNil(this.min) ? Number.NEGATIVE_INFINITY : this.min
          //let max = R.isNil(this.max) ? Number.POSITIVE_INFINITY : this.max
          let fValueAsPct = this.decimalToPercent(fValueRaw)
          return fValueAsPct
        },
        set(newValue) {
          let fValueRaw = typeof newValue === 'number' ? newValue : parseFloat(newValue)
          if (isNaN(fValueRaw)) {
            this.$emit('update:modelValue', null)
          } else {
            let min = R.isNil(this.min) ? Number.NEGATIVE_INFINITY : this.min
            let max = R.isNil(this.max) ? Number.POSITIVE_INFINITY : this.max
            let fValueClamped = R.clamp(min, max, fValueRaw)
            let fValue = fValueClamped/100
            
            if (fValueClamped !== fValueRaw) {
              this.$refs.pct.value = fValueClamped
            }
            this.$emit('update:modelValue', fValue)
          }
  
        }
      },
      modelValueAsCurrency () {
        const formatter = new Intl.NumberFormat(this.currencyLocale, {style:'currency', currency: this.currency})
        return isNaN(this.modelValueAsNumber) ? '' : formatter.format(this.modelValueAsNumber )
      },
      
      modelValueAsSwitch () {
        let valueForCheckbox = this.modelValue
        if (!this.isBooleanType) return !!this.modelValue
        if (typeof valueForCheckbox === 'boolean') {
          return valueForCheckbox
        } else {
          if (valueForCheckbox == this.switchOnValue) {
            return true
          } else if (valueForCheckbox == this.switchOffValue) {
            return false
          } else {
            console.warn(`The value passed for the switch "${valueForCheckbox}" does not conform to either the specified "On" value of "${this.switchOnValue}" or the "Off" value "${this.switchOffValue}".\nThe off state will be assumed.\nRecommend that you check your data bindings!`)
            return false
          }
        }
      },
  
      // For when a Mutil select is specified
      unusedOptions () {
        let a = this.modelValue
        console.warn('Not yet implemented a generic filter on items already used ...')
        return a 
        let aPropEqualTests = this.optionValuePropertyList.map(testPropName => R.eqProps(testPropName)) // The array is now filled with 2 parameter test functions
        let fnOptionIsInModelValue = (item) => a.findIndex(itemInModel => R.eqProps( this.optionValueProperty,item, itemInModel)) //this.modelValue
        return R.compose(
  
          R.filter(fnOptionIsInModelValue)
        )(this.options)
  
      },
      suppressMainLabel () {
        let aTypeToSuppress = ['switch']
        return aTypeToSuppress.includes(this.type)
      },
      haveSelectedItemToAddToGroup () {
        return !R.isNil(this.selectedItemToAddGroup)
      }
    
    },
    methods: {
      
      decimalToPercent(value) {
        const fnGetPlaces = (num) => {
          let numAsText = num.toString()
          let idxDecimalPoint = numAsText.indexOf('.')
          return idxDecimalPoint > - 1 ? numAsText.length - idxDecimalPoint - 1 : 0
        }
        let srcPrecision = fnGetPlaces(value)
        let valueAsPercent = value * 100
        let trgPrecision = fnGetPlaces(valueAsPercent)
        if (trgPrecision > srcPrecision + 2){
          return valueAsPercent.toFixed(srcPrecision)
        } else {
          return valueAsPercent.toString()
        }
      },
      parseCurrencyText (newValue) {
        let fValue
        if (typeof newValue === 'string') {
          let hadSign = false
          let hadDecimal = false
          let reIsDecimal = /[0-9]/
          let chars = newValue.split('')
          let stripped = ''// newValue.replace(reAllowed,'')
          chars.forEach(c =>{
            if (reIsDecimal.test(c)){
              stripped +=c
            } else {
              if (!hadSign && (c == '-' || c == '+')) {
                stripped += c
                hadSign = true
              }
              if (!hadDecimal && c == '.') {
                stripped += c
                hadDecimal = true
              }
            }
          })
          fValue = parseFloat(stripped)
        } else {
          fValue = parsFloat( newValue)
        }
        return fValue
      },
      updateCheckBoxes() {
        // For a type checkboxes the value to emit is an array of values from those checkboxes that are TRUE
        let aChecks = this.$refs.checkbox
        let aSelectedChecks = aChecks.filter(checkbox => checkbox.checked)
        let modelArray = aSelectedChecks.map(checkbox => checkbox.value)
        this.$emit('update:modelValue', modelArray)
  
      },
      onChange(ev) {
        let value = ev.target.value
        
        if (this.isDateType) {
          let isoDateTextForSharePoint = this.formatControlDateToSharePointIso(value)
          value = isoDateTextForSharePoint
        } else if (this.isNumberType) {
          let valueAsNumber 
          if (this.isCurrencyType) {
            valueAsNumber = this.parseCurrencyText(value)
          } else {
            valueAsNumber = this.valueAsNumber
          }
          
          let maxClamp = this.$attrs.max ? parseFloat(this.$attrs.max) : undefined
          let minClamp = this.$attrs.max ? parseFloat(this.$attrs.min) : undefined
          if (!isNaN(maxClamp) && valueAsNumber > maxClamp ) {
            value =  maxClamp.toString() // Because this is how the updateModel handler would EXPECT to get the value, a string and not a number!
            ev.target.value = value
          }
  
          if (!isNaN(minClamp) && valueAsNumber < minClamp ) {
            value =  minClamp.toString() // Because this is how the updateModel handler would EXPECT to get the value, a string and not a number!
            ev.target.value = value
          }
          if (isNaN(valueAsNumber)) {
            value = null // An Empty string value (or other non number) will cause an error in SharePoint attempting to convert data to "Edm.Double" so better pass NULL when duff value received!
          } else {
            value = valueAsNumber
          }
          // SELECT CONTROLS HAVE THEIR OWN UPDATE LOGIC!
        // }  else if (this.type === 'select' && this.optionBindObject) {
        //   // Get the underlying object 
        //   // value = this.options[ev.target.selectedIndex - 1] // The -1 is because we have generates a (Choose ...) entry
  
        //   value = this.options[value]
        } else if (this.isBooleanType) {
          
          value = !!ev.target.checked
  
          if (this.type === 'switch') {
            if (value) {
              value = this.switchOnValue
            } else {
              value = this.switchOffValue
            }
            this.$emit('checked',value)
          }
        }
        
        this.$emit('update:modelValue', value)
        
  
        
      },
      getOptionIndex(opt) { 
        // need to determine method to compare passed option to items in list of options
        if (opt === null) return undefined;
        
        if (!this.options.map || this.options.length< 1)  return undefined;

        if (typeof opt === 'object'){
          let testFunction
          if (this.optionValueProperty) {
            console.log('Here in the object specified value')
            testFunction  = R.eqProps(this.optionValueProperty) // Current to a 2 parameter test function
          } else if (this.optionTextProperty){
            testFunction  = R.eqProps(this.optionTextProperty) // Current to a 2 parameter test function
          } else if (this.optionFormatterFunction) {
            testFunction = (a,b) => this.optionFormatterFunction(a) == this.optionFormatterFunction(b)
          } else {
            testFunction = R.equals
          }      
          return this.options.findIndex(itm => testFunction(itm,opt))
        } else {
          // because the "opt" is a simple value we need to reduce the options to a simple value too
          let list = []
          if (this.optionValueProperty) {
            console.log('Here in the simple specified value')
            list = this.options.map(itm => itm[this.optionValueProperty])
            
          } else if (this.optionTextProperty){
            list = this.options.map(itm => itm[this.optionTextProperty])
            
          } else if (this.optionFormatterFunction) {
            list = this.options.map(itm => this.optionFormatterFunction(itm))
            
          } else {
            list = this.options.slice(0)
          } 
          //console.log(list)
          return list.indexOf(opt)
        }
      }, // The objects are proxies and not going to be ===, but Ramda whereEq a good test
      getBestOptionValue(opt) {
        if (typeof opt === 'object') {
          
         
          if (this.optionValueProperty) {
            return opt[this.optionValueProperty]
          } else {
            return opt.code ?? opt.key ?? opt.Id ?? opt.id ?? opt.name ?? opt.toString()
          }
        } else {
          return opt
        }
      },
      getCheckedOptionValue (opt) {
        
        if (this.modelValue instanceof Array) {
          let value = this.getBestOptionValue(opt)
          return this.modelValue.includes(value)
        } else {
          return !!this.modelValue // cannot be checked list, but if a simple property just use its "truthiness"
        }
        
      },
      getBestOptionTextStandard(opt) {
        if (typeof opt === 'object') {
          if (this.optionTextProperty) {
            return opt[this.optionTextProperty]
          } else {
            return opt.caption ?? opt.title ?? opt.Title ?? opt.name ?? opt.text ?? opt.display ?? opt.toString()
          }        
        } else {
          return opt
        }
  
      },
      getBestOptionText(opt) {
        if (typeof this.optionFormatterFunction === 'function') {
          return this.optionFormatterFunction.call(this, opt)
        } else {
          if(typeof opt === 'number') {
            return this.getBestOptionTextStandard(this.options[opt])
          } else {
            
            return this.getBestOptionTextStandard(opt)
          }
        }          
      },
      localeDateFromIso(isoText) {
        const reIsoDate = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2}).(\d{1,3}){0,1}/
        if (R.isNil(isoText) || !reIsoDate.test(isoText)) return null;
      
        let parts = reIsoDate.exec(isoText)
        .slice(1,8)
        .map(p=>parseInt(p,10))
        .filter(p=> p >=0)
        parts[1]-- // decrement month to 0 to 11 basis
        let timeCode = Date.UTC(...parts)
        return new Date(timeCode)
      },
      formatDateInputControl(isoText) {
        let dtm = this.localeDateFromIso(isoText)
        let mmt= moment(dtm)
        return  mmt.isValid() ? mmt.format('YYYY-MM-DD') : ''
      },
      formatDateTimeInputControl(isoText) {
        let dtm = this.localeDateFromIso(isoText)
        let mmt= moment(dtm)
        return  mmt.isValid() ? mmt.format('YYYY-MM-DDT:HH:mm') : ''
      },
      formatControlDateToSharePointIso(controlText) {
        //.toDate().toISOString().replace(/\.\d*/,''
        if (controlText==='') return null;
        const TICKS_PER_MIN = 1000 * 60
        let dtm = new Date(controlText)
        let offsetMinutes = dtm.getTimezoneOffset()
        let timeCodeLocal = dtm.valueOf()
        let timeCodeUtc = timeCodeLocal + (TICKS_PER_MIN * offsetMinutes)
  
        return (new Date(timeCodeUtc)).toISOString().replace(/\.\d{0,3}/,'')
      },
      getOptionStyle (option) {
        let oStyle ={}
        let backgroundProperty = this.optionBackgroundColourProperty
        let colourProperty = this.optionForegroundColourProperty
        if (backgroundProperty) {
          if (option.hasOwnProperty(backgroundProperty)){
            oStyle['background-color'] = option[backgroundProperty]
          }
        }
        if (colourProperty) {
          if (option.hasOwnProperty(colourProperty)){
            oStyle['color'] = option[colourProperty]
          }
        }
        return oStyle
      },
      removeFromModelArrayAtIndex(idx) {
        console.assert(this.modelValue instanceof Array, 'An Array was expected ...')
        
        this.$emit('update:modelValue', R.remove(idx,1,this.modelValue))
      },
      appendToModelArray() {
        console.assert(this.modelValue instanceof Array, 'An Array was expected ...')
        let itemToAdd = this.options[this.selectedItemToAddGroup]
        this.$emit('update:modelValue', R.append(itemToAdd,this.modelValue))
      }
  
    },
    mounted(){ 
      this.localId = `${Math.random().toString().substring(2)}` 
      if (this.type === 'checkboxes' && !(this.modelValue instanceof  Array)) {
        this.$emit('update:modelValue', [])
  
      }
      // const aLookupTypes = ['lookup','lookupMulti','user','userMulti']
      // if (aLookupTypes.includes( this.type )) {
      //   // Need to populate the 
  
      // }
      if (this.type === 'select' ) {
        let idx = this.getOptionIndex(this.modelValue)
        if (idx >-1){
          this.boundSelectIndex = idx
        }
      }
    },
    template:`<div :class="$attrs.class" :style="$attrs.style">
    <label v-if="!suppressMainLabel" :for="renderedControlId" :class="effectiveLabelClass">{{label}}</label>
    <div class="input-group" v-if="useGroup">
      <span v-if="!suppressPrefixIcon" :id="'typeIcon-' + localId" class="input-group-text"><i :class="typeIconClass"></i></span>
      <textarea v-if="type === 'textarea'"
        :readonly="readonly"
        :value="modelValue"
        class="form-control"
        @input="onChange($event)"
        :id="renderedControlId"
        :rows="rows"
        :placeholder="placeholder"
        >
      </textarea>
      <select v-else-if="type === 'select'"
        class="form-select" 
        :readonly="readonly"
        v-model="boundSelectIndex"
        :id="renderedControlId"
        :style="getActiveOptionStyle"
      >
        <option disabled >(Choose ...)</option>
        <option v-for="(opt,idx) in options" :key="'opt-' + localId + '-' + idx" :value="idx" :style="getOptionStyle(opt)">
        {{getBestOptionText(opt)}}
        </option>
      </select>
  
  
      <input v-else-if="type === 'date'"
        :readonly="readonly"
        :type="type"
        :value="formatDateInputControl(modelValue)"
        class="form-control"
        @change="onChange($event)"
        :id="renderedControlId"
        
      >
  
      <input v-else-if="type === 'datetime-local'"
        :type="type"
        :readonly="readonly"
        :value="formatDateTimeInputControl(modelValue)"
        class="form-control"
        @change="onChange($event)"
        :id="renderedControlId"
        
      >
  
      <input v-else-if="type === 'text'"
        :type="type"
        :readonly="readonly"
        class="form-control"
        :value="modelValue"
        @input="onChange($event)"
        :maxlength="maxlength"
        :id="renderedControlId"
        :placeholder="placeholder"
        :list="renderedControlId + '-datalist'"
      >
  
      <input v-else-if="type === 'number'"
        type="number"
        :readonly="readonly"
        class="form-control"
        :id="renderedControlId"
        :value="modelValueAsNumber"
        @input="modelValueAsNumber = $event.target.value"
        :min="min"
        :max="max"
        :step="step"
        ref="num"
      >
      <input v-else-if="type === 'percent'"
        type="number"
        :readonly="readonly"
        class="form-control"
        :id="renderedControlId"
        v-model.number="modelValueAsPercentage"
        ref="pct"
      >
  
      <input v-else-if="type === 'currency'"
        type="text"
        :readonly="readonly"
        class="form-control"
        @change="onChange($event)"
        :value="modelValueAsCurrency"
        :id="renderedControlId"
      >
  
      <input v-else 
        :type="type" 
        :readonly="readonly"
        class="form-control"
        :value="modelValue"
        @input="onChange($event)"
        :id="renderedControlId"
        
      >
      <span v-if="required" class="input-group-text">
        <span v-if="haveValue"><i :class="validIconClass"></i></span>
        <span  v-if="!haveValue"><i :class="invalidIconClass"></i></span>
      </span>
      <slot name="groupsuffix"></slot>
  
      <datalist :id="renderedControlId + '-datalist'" v-if="options?.length > 0 && type==='text'">
        <option v-for="(opt,idx) in options" :key="'opt-' + localId + '-' + idx" :value="getBestOptionValue(opt)">{{getBestOptionText(opt)}}</option>
      </datalist>
    
    </div>
  
  
    <div v-if="!useGroup">
  
      <div v-if="type === 'checkboxes'" class="form-check form-check-inline" :id="renderedControlId">
        <template v-for="(check, idx) in options"  :key="renderedControlId + '-' + idx" >
        <div  class="form-check form-check-inline">
          <input type="checkbox" ref="checkbox"
            class="form-check-input"  :id="renderedControlId + '-' + idx" 
            :value="getBestOptionValue(check)" 
            :checked="getCheckedOptionValue(check)"
            @change="updateCheckBoxes"
            >
          <label class="form-check-label" :for="renderedControlId + '-' + idx">{{getBestOptionText(check)}}</label>
        </div>
        </template>
      </div>
  
      <div v-else-if="type === 'radio'"  :id="renderedControlId">
  
        <template v-for="(check, idx) in options"  :key="renderedControlId + '-' + idx" >
        <div  class="form-check form-check-inline">
          <input class="form-check-input" type="radio" :id="renderedControlId + '-' + idx" :value="getBestOptionValue(check)" :checked="getCheckedOptionValue(check)">
          <label class="form-check-label" :for="renderedControlId + '-' + idx">{{getBestOptionText(check)}}</label>
        </div>
        </template>    
      </div>
  
      <div v-else-if="type === 'select'" :class="$attrs.class">
        <div class="w-100 border border-1 p-1">
          <TransitionGroup name="simple-fade">
          <span v-for="(item,idx) in unusedOptions" :key="renderedControlId + '-' + idx + item.Id || '-'"
            class="d-inline-block badge  me-2 mb-1 text-bg-info border border-primary"
          ><small>
            {{getBestOptionText(item)}}&nbsp;<button class="btn nav-link text-dark" role="button" @click="removeFromModelArrayAtIndex(idx)"><i class="fas fa-trash"></i></button>
          </small></span></TransitionGroup>
        </div>
        <div class="mt-1 input-group">
          <select
            v-model="selectedItemToAddGroup"
            class="form-select  form-select-sm" 
            :id="renderedControlId"
          >
            
            <option v-for="(opt,idx) in options" :key="'opt-' + localId + '-' + idx" :value="idx" :style="getOptionStyle(opt)">
            {{getBestOptionText(opt)}}
            </option>
          </select>
          <button @click="appendToModelArray" :disabled="!haveSelectedItemToAddToGroup" class="btn btn-small btn-primary">+</button>
        </div>
      </div>
      
      <div v-else-if="type === 'switch'" class="form-check form-switch">
        <input class="form-check-input" type="checkbox" role="switch"
          :id="renderedControlId" :checked="modelValueAsSwitch" @change="onChange($event)"
        >
        <label class="form-check-label" :for="renderedControlId">{{label}}</label>      
      </div>
  
  
  
    </div>
    
    <slot></slot>
  </div>`
   }

}






// // Automatically install if Vue has been added to the global scope.
// if (typeof window !== 'undefined' && window.Vue) {

//   window.Vue.use(dhgBootstrapWrapperPack)
// }