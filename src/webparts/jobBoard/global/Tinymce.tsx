import * as React from 'react';
import { Editor } from '@tinymce/tinymce-react';
import { EventHandler } from 'react';

export interface ITinymceProps {
  onChange  : EventHandler<any>;
  defaultValue? : string;
  limit? : number;
}

export class Tinymce extends React.Component<ITinymceProps, {}> {
  constructor(props: ITinymceProps) {
    super(props);
  }

  public render() {
    return (
      <Editor
        initialValue ={this.props.defaultValue}
        init={{
          menubar:false,
          statusbar: true,
          height : 300,
          plugins: "wordcount",
          wordcount_cleanregex: /[0-9.(),;:!?%#$?\x27\x22_+=\\\/\-]*/g,
          toolbar: 'undo redo | bold italic underline strikethrough superscript subscript | alignleft aligncenter alignright'
        }}
        onChange={this.props.onChange}
      />
    );
  }
}

