import * as React from 'react';
import { Editor } from '@tinymce/tinymce-react';
import { EventHandler } from 'react';

export interface ITinymceProps {
  onChange  : EventHandler<any>;
}

export class Tinymce extends React.Component<ITinymceProps, {}> {
  constructor(props: ITinymceProps) {
    super(props);
  }

    // Returns text statistics for the specified editor by id
  /*private _getStats(id) {
    //let body = tinymce.get(id).getBody()
    //let text = tinymce.trim(body.innerText || body.textContent);
    return {
        chars: text.length,
        words: text.split(/[\w\u2019\'-]+/).length
    };
  }*/

  public render() {
    return (
      <Editor
        initialValue=""
        init={{
          menubar:false,
          statusbar: false,
          height : 300,
          toolbar: 'undo redo | bold italic underline strikethrough superscript subscript | alignleft aligncenter alignright'
        }}
        onChange={this.props.onChange}
      />
    );
  }
}

