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

  public render() {
    return (
      <Editor
        initialValue="<p>This is the initial content of the editor</p>"
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

