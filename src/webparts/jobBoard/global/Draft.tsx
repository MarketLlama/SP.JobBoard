import * as React from 'react';
import styles from './../components/JobBoard.module.scss';
import { EditorState, convertToRaw,  ContentState} from 'draft-js';
import { Editor } from 'react-draft-wysiwyg';
import 'react-draft-wysiwyg/dist/react-draft-wysiwyg.css';
import { stateFromHTML } from 'draft-js-import-html';
import { stateToHTML } from 'draft-js-export-html';

export interface DraftProps {
  onChange  : Function;
  defaultValue? : string;
  limit? : number;
}

export interface DraftState {
  editorState : EditorState;
  wordCount? : number;
}

class Draft extends React.Component<DraftProps, DraftState> {

  constructor(props: DraftProps) {
    super(props);
    if(this.props.defaultValue){
      let contentState = stateFromHTML(this.props.defaultValue);
      this.state = {
        editorState : EditorState.createWithContent(contentState)
      };
    } else {
      this.state = {
        editorState : EditorState.createEmpty()
      };
    }
  }
  public render() {
    return (<div>
    <div className={styles.draftEditor}>
      <Editor
        editorState={this.state.editorState}
        toolbarClassName={styles.toolbar}
        wrapperClassName={styles.wrapper}
        editorClassName={styles.editor}
        onEditorStateChange={this._onEditorStateChange}
      />
    </div>
      <div>
        <span>Word Count : {this.state.wordCount}</span>
      </div>
    </div>);
  }

  private _onEditorStateChange = (value :EditorState) =>{
    this.setState({
      editorState : value
    });
    let contentState = value.getCurrentContent();
    this._getWordCount();
    let html = stateToHTML(contentState);
    this.props.onChange(html);
  }

  private _getWordCount = () : void => {
    const plainText = this.state.editorState.getCurrentContent().getPlainText('');
    const regex = /(?:\r\n|\r|\n)/g;  // new line, carriage return, line feed
    const cleanString = plainText.replace(regex, ' ').trim(); // replace above characters w/ space
    const wordArray = cleanString.match(/\S+/g);  // matches words according to whitespace
    this.setState({
      wordCount :wordArray ? wordArray.length : 0
    });
  }
}

export default Draft;
