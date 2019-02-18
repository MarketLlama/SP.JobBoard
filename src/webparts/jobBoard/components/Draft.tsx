import * as React from 'react';
import styles from './../components/JobBoard.module.scss';
import { EditorState, convertToRaw,  ContentState} from 'draft-js';
import { Editor } from 'react-draft-wysiwyg';
import 'react-draft-wysiwyg/dist/react-draft-wysiwyg.css';

export interface IDraftHelperProps {
  defaultValue? : string;
  limit? : number;
}

export interface IDraftHelperState {
  editorState : EditorState;
  wordCount? : number;
}

export class DraftHelper extends React.Component<IDraftHelperProps, IDraftHelperState> {

  constructor(props: IDraftHelperProps) {
    super(props);

    this.state = {
      editorState : EditorState.createEmpty()
    };
  }
   render() {
    return (<div>
    <div className={styles.draftEditor}>
      <Editor
        editorState={this.state.editorState}
        toolbarClassName={styles.toolbar}
        wrapperClassName={styles.wrapper}
        editorClassName={styles.editor}
        onEditorStateChange={(editorState) => { this.setState({editorState})}}
      />
    </div>
      <div>
        <span>Word Count : {this.state.wordCount}</span>
      </div>
    </div>);
  }
}

