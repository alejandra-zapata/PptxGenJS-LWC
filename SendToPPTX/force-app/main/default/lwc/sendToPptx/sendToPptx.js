import { LightningElement } from 'lwc';
import { loadScript } from 'lightning/platformResourceLoader';
import p from '@salesforce/resourceUrl/pptxgenBundle';

// This class is trying to load the JS library as a static resource and creating a sample PPTX presentation with text as the user clicks on 'Send to Pptx'
export default class SendToPptx extends LightningElement {
    scriptsLoaded = false;

    connectedCallback() {
        if(this.scriptsLoaded) {
            return;
        }

        loadScript(this, p)
            .then(() => {
                this.scriptsLoaded = true;
                console.log('Scripts are loaded successfully.');
            })
            .catch(error => {
                console.log('Error loading the scripts: ' + p);
            })
    }

    handleSend() {
        try{
            let pptx = new PptxGenJS();
            let slide = pptx.addSlide();
            let opts = {
                x: 0.0,
                y: 0.25,
                w: '100%',
                h: 1.5,
                align: 'center',
                fontSize: 24,
                color: '0088CC',
                fill: 'F1F1F1'
            };
            slide.addText(
                'From LWC',
                opts
            );
            pptx.writeFile();
        }
        catch(error){
            console.log('Error writing presentation.***** ' + error.name + ' **** ' + error.message + ' **** ' + error.stack);
        }
    }
}