
/**
 * @description
 * This component is a Vue JS (V3+) component that creates a set of visual controls to allow a user to browse a a SharePoint site document 
 * library for a file using the OneDrive File Picker SDK (v8) relased in June 2022
 * @requires vue (global)
 * @requires Bootstrap 5.3+ (global)
 * @fontAwesome 5 (global)
 * 
 */

// Define Internal variables here
const _getOptions = { headers: { Accept: 'application/json; odata=nometadata' } }

async function getSiteDrives (siteUrl) {
    let url =`${siteUrl}/_api/v2.1/drives`
    try {
    let resp = await fetch(url,_getOptions)
    let data = await resp.json()
    return data.value
    } catch (ex) {
      console.error(`Failed in vueSpFileBrowser::getSiteDrives\nURL:\n${url}\n${ex.message}`)
      throw ex
    }
  }
  async function getSiteDriveFileDetails (site, drive, file) {
    if (typeof site === 'string') {
      if (!site.endsWith('/')) site += '/'
    }
    let url = `${site}_api/v2.1/drives/${drive}/items/${file}`
    let resp = await fetch (url, _getOptions)
    let data = await resp.json()
    return data
  }

/**
 * 
 * @param {String} webUrl - Full path to a resource
 * @returns decoded relative path to server
 */
export function relativePathFromUrl(webUrl){
  const re = /^(https?:\/\/)([^\/]+)(\/[^?#]*)?/i
  if (re.test(webUrl)){
    let mCol = re.exec(webUrl)
    return decodeURIComponent(mCol[3])
  } else {
    return  ''
  }  
}
function parentFolderFromRelativePath(relativePath){
  const aParts = relativePath.split('/')
  if (aParts.length){
    let fileName = aParts.pop() // the array is mutated!
    return aParts.join('/')
  } else {
    return ''
  }

}

const vueSpFileBrowser = {
    name: 'vueSpFileBrowser',
    inheritAttrs: true,
    data() {
      
      return {

        ignoreWatches: true,
        popupWindow: null, // Will be populated after a SHOW,
        randomId: Math.random(32).toString().substring(2),
        fnMessageListener: null,
        drives:[],
        selectedSiteUrl: '',
        selectedDriveId: null,
        lastDriveId: null,
        selectedFileId: null,
        lastFileId: null,
        filePicked: null,
        filePickedFolder: null
      }
    },
    props: {
      // Graph file requirements
      siteUrl: {type: String, required: true },
      driveId: {type: String, required: false },
      fileId: {type: String, required: false },

      showSiteControl: {type: Boolean, required: false, default: false},
      siteLabel: {type: String, required: false, default: 'Site:'},
      showDriveControl: {type:Boolean, required: false, default: true},
      showDriveDescription: {type:Boolean, required: false, default: true},
      showDriveFileCount: {type:Boolean, required: false, default: true},
      driveLabel:{type: String, required: false, default: 'Library:'},

      
      title: {type: String, required: false, default: 'Select a file ..'},
      fileNamePlaceholder: {type: String, required: false, default: 'Selected file ..'},

      buttonCaption: {type: String, required: false, default: 'File ...,'},
      buttonIcon: {type: String, required: false, default: 'fa-folder-open'},
      buttonClass: {type: String, required: false, default: 'btn-primary'},
      
      
      typeFilters: {type:String, required: false, default: ''},
      disabled: {type:Boolean, required: false, default: false},
      openTarget: {type:String, required: false, default: '_blank'},
      indicateListening: {type:Boolean, required: false, default: false},
    },
    async mounted(){
      const {fileId,driveId,siteUrl}= this
      if (fileId) this.selectedFileId = fileId;
      if (driveId) this.selectedDriveId = driveId;
      if (siteUrl) this.selectedSiteUrl = siteUrl;
      let aPromises =[]
      aPromises.push( this.loadDrivesForSite(siteUrl))
      if (fileId){
        // console.log(`getting graph for:-\n${siteUrl}\n${driveId}\n${fileId}`)
        aPromises.push (getSiteDriveFileDetails(siteUrl,driveId,fileId))
      }
      
      let aResolved = await Promise.allSettled(aPromises)
      // console.log(aResolved)
      if (aResolved[1]){
        // we have teh "picked" file details
        let graphFile = aResolved[1].status === 'fulfilled' ? aResolved[1].value : null
        if (graphFile) {
          // console.log(graphFile)
          this.filePicked = graphFile
        } else {
          console.warn('Failed to get the expected graph file specification!')
        }
      }
      window.setTimeout(()=> this.ignoreWatches=false, 350)
    },
    watch:{
      popupWindow(newWindow,oldValue){
        if (newWindow) {
          let hInterval = window.setInterval(()=>{if (newWindow.closed){
            window.clearInterval(hInterval)
            if (this.fnMessageListener) this.stopListening();
          }},500)
        }
      },
      filePicked(newValue,oldValue){
        if (!this.ignoreWatches){
          this.$emit('filePicked', newValue)
        }

        
        if (!newValue){
          this.filePickedFolder = null
        } else {
          getSiteDriveFileDetails(this.selectedSiteUrl, newValue.parentReference?.driveId, newValue?.parentReference.id)
            .then (folder=>{
              this.filePickedFolder = folder
            })
            .catch(error=>{
              let msg = `Error fetching parent folder details to picked file.\nError: ${error}`
              console.error(msg)
              this.$emit('error', msg)
              this.filePickedFolder = null
            })
        }
        
      },
      selectedSiteUrl: {
        
        handler(newUrl,oldUrl){
          if (this.ignoreWatches) return;
          if (!!newUrl){
            this.loadDriveForSite(newUrl)
              .then(()=>{
                if (this.drives.length){
                  this.selectedDriveId= this.drives[0].id
                }
              })
              .catch((error)=>{
                this.drives = []
              })
            
          } else {
            this.clearAllSelected()
          }
        }
      },
      selectedDriveId:{
        handler(newDriveId, oldDriveId){
          if (this.ignoreWatches) return;
          let currentFile = this.selectedFileId
          if (currentFile) {
            this.pickFile = null
          }
        },
        immediate: false
      }

    },
    computed: {
      dialogId(){return `dialog${this.randomId}`},
      siteCtrlId(){return `siteCtrl${this.randomId}`},
      driveCtrlId(){return `driveCtrl${this.randomId}`},
      fileCtrlId(){return `fileCtrl${this.randomId}`},
      haveDriveDescription(){
        if (!this.showDriveDescription) return false; // Suppressed by attribute override
        let drive = this.drives.find(drv => drv.id === this.selectedDriveId)
        return drive && !!drive.description
      },
      selectedDriveDescription(){
      
        let drive = this.drives.find(drv => drv.id === this.selectedDriveId)
        return drive?.description
      },
      filePickedName(){
        let fileName='File Name ...'
        if (this.filePicked){
          fileName = this.filePicked?.name
        }
        return fileName
      },
      searchButtonIconClasses(){
        let aClasses = this.buttonIcon.replace(/\s+/g,'#').split('#')
        if (!aClasses.includes('fas')) aClasses.unshift('fas');
        return aClasses
      },
      isPickDisabled(){
        return this.disabled || !this.selectedDriveId
      },
      haveFilePicked(){ return !!this.filePicked},

      pickerUrl: function () {
        
        const siteBase = this.selectedSiteUrl
        const siteRelative = relativePathFromUrl(siteBase)

        const filePickedFolder = this.filePickedFolder?.webUrl

        let idParam = `${siteRelative}/Shared Documents` // A reasonable stab at a default id
        if (filePickedFolder) {
          
          idParam = encodeURIComponent(relativePathFromUrl(filePickedFolder))
        } else if (this.selectedDriveId){

          let drive = this.drives.find(drv => drv.id === this.selectedDriveId)
          if (drive) {
            idParam = encodeURIComponent(relativePathFromUrl(drive.webUrl))
          }
        }

        let urlTypeFilters = this.typeFilters.trim() ? `&typeFilter=${this.typeFilters}` :''
        return `${siteBase}/_layouts/15/OneDrive.aspx?p=2&id=${idParam}${urlTypeFilters}`
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
      async loadDrivesForSite(siteUrl){
        try {
          let aDrives = await getSiteDrives(siteUrl)
          this.drives = aDrives

        } catch (error) {
          let msg = `Error loading drives from site.\nError: ${error}`
          console.error(msg)
          this.$emit('error', msg)
          throw error
        }
        
      },
      clearAllSelected(){
        
        if (this.showDriveControl) this.selectedDriveId = null
        if (this.showSiteControl) {this.selectedSiteUrl = ''} else {this.selectedSiteUrl = this.siteUrl}
        this.selectedFileId = null
        this.filePicked =  null
      },
      openSelected(){
        let url = this?.filePicked?.webUrl
        if (url) window.open(url,this.openTarget)
      },
      setFile(graphFile){
        const fileId = graphFile?.id
        const driveId = graphFile?.parentReference?.driveId
        const siteUrl = graphFile?.parentReference?.sharepointIds?.siteUrl
        if (fileId && driveId && siteUrl){
          this.selectedSiteUrl = siteUrl
          this.selectedDriveId = driveId
          this.selectedFileId = fileId
        }
      },
      receiveMessage: function (evt) {
        if (!evt.isTrusted) return ;
        const PICKER_EVENT = '[OneDrive-FromPicker]'
        
        try {
          let eventData = evt.data

          if (('' + eventData).startsWith(PICKER_EVENT)) {
            
            let pickedData = JSON.parse(eventData.substring(PICKER_EVENT.length))
            switch (pickedData.type) {
              case 'cancel':
                
                this.cancelPopup()
                break
              case 'success':
                
                let firstPickerFile = pickedData.items[0] 
                const driveId = firstPickerFile?.parentReference?.driveId
                const fileId = firstPickerFile.id
                const siteUrl = this.selectedSiteUrl
                getSiteDriveFileDetails(siteUrl,driveId,fileId)
                .then(graphFile =>{
                  this.setFile(graphFile)
                  this.filePicked = graphFile
                  this.$emit('filePicked', graphFile)
                })
                .catch(error=>{
                  let msg = `Unexpected error attempting to get Graph details following a picker selection.\n${error}\nSite: ${siteUrl}\nDrive: ${driveId}\nFile ID: ${fileId}`
                  console.error(msg)
                  console.log('firstPickerFile')
                  console.log(firstPickerFile)
                  this.$emit('error',error)
                })

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
          this.showPickerWindow()
          
          this.$emit('startedListening', this.pickerUrl)
          
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
        
        window.addEventListener('message', fnMessageListener, false)
      },
      stopListening: function () {
        if (this.fnMessageListener) {
          
          try {
            window.removeEventListener('message', this.fnMessageListener)
            
            this.fnMessageListener = null
            this.$emit('stoppedListening')
          } catch (ex) {

            console.error(`Failed to remove the message listener`)
          }
        }

      },
      cancelPopup: function() {
        this.stopListening()
        this.cancelPopupWindow()
      },
    
      cancelPopupWindow: function () {
        if (this.popupWindow) {
          try{
            this.popupWindow.close()
          } catch (ex) {
            console.error(`Failed calling close method on popupWindow.\nError: ${ex instanceof Error ? ex.message : ex?.toString()}`)
          }
          this.popupWindow = null
        } else {
          console.warn('Popup Picker is not active!')
        }
      }
    },
    emits: ['filePicked', 'startedListening','stoppedListening','error'],
    expose: ['showPicker', 'cancelPopup', 'startListening'],
    template: `
<div class="position-relative">
  <div v-if="indicateListening && fnMessageListener" class="position-absolute top-50 start-100 translate-middle" title="Waiting for selection ..."><i class="fas fa-pulse fa-wave-square"></i></div>
  <div class="w-100" v-if="showSiteControl">
    <div class="input-group mb-2">
      <span class="input-group-text" ><label :for="siteCtrlId">{{siteLabel}}</label></span>
      <input class="form-control" v-model.lazy.trim="selectedSiteUrl" :id="siteCtrlId" >
        
    </div>
  </div>

  <div class="w-100" v-if="showDriveControl">
    <div class="input-group mb-2">
      <span class="input-group-text" ><label :for="driveCtrlId">{{driveLabel}}</label></span>
      <select class="form-select" v-model="selectedDriveId" :id="driveCtrlId">
        <option selected>Choose...</option>
        <option v-for="library in drives" :key="library.id" :value="library.id">{{ library.name }}<span v-if="showDriveFileCount"> ({{ library.quota.fileCount }} files)</span></option>
      </select>
    </div>
    <div v-if="haveDriveDescription">{{selectedDriveDescription}}</div>
  </div>

  <div class="w-100">
    <div class="input-group mb-2">
      <span class="input-group-text"  >File:</span>
      <input type="text" class="form-control"
        :id="fileCtrlId"
        disabled :placeholder="fileNamePlaceholder"
        :aria-label="fileNamePlaceholder"
        
        :value="filePicked?.name">
      <button
        class="btn" :class="derivedButtonClass"
        @click.prevent="showPicker"
        :disabled="isPickDisabled"
      >
          <i v-if="buttonIcon" :class="searchButtonIconClasses"></i>
          {{buttonCaption}}
      </button>
        
      <button class="btn btn-outline-warning" @click="clearAllSelected" >Clear</button>     
      <button class="btn btn-outline-secondary" @click="openSelected" :disabled="!haveFilePicked">View (Browser)</button>     
    </div>
  </div>  
</div>
`
  }

  export default vueSpFileBrowser