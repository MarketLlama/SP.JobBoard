import * as React from 'react';
import ReactQuill, { Quill } from 'react-quill';
import 'react-quill/dist/quill.snow.css';

export interface QuillServiceProps {
  onChange  : Function;
  defaultValue? : string;
  limit? : number;
}

export interface QuillServiceState {
 editorValue : string;
 wordCount? : number;
}

class QuillService extends React.Component<QuillServiceProps, QuillServiceState> {
  constructor(props) {
    super(props)
    if(this.props.defaultValue){
      this.state = ({
        editorValue : this.props.defaultValue
      })
    } else {
      this.state = { editorValue: '' }
    }
  }


  modules = {
    toolbar: [
      [{ 'header': '1'}, {'header': '2'}, { 'font': [] }],
      [{ 'size': ['14px', '18px', '24px'] }],
      ['bold', 'italic', 'underline', 'strike', 'blockquote'],
      [{'list': 'ordered'}, {'list': 'bullet'},
       {'indent': '-1'}, {'indent': '+1'}],
      ['clean']
    ],
    clipboard: {
      matchVisual: true,
    }
  };

  formats = [
    'header', 'font', 'size',
    'bold', 'italic', 'underline', 'strike', 'blockquote',
    'list', 'bullet', 'indent'
  ];

  private _handleChange = (editorValue) => {
    this.setState({ editorValue})
    this.props.onChange(editorValue);
    this._getWordCount();
  }

  render() {

    return (<div>
      <div >
        <ReactQuill value={this.state.editorValue}
          modules={this.modules}
          formats={this.formats}
          onChange={this._handleChange}
        />
      </div>
        <div>
          <span><strong>Word Count : </strong>{this.state.wordCount}</span>
        </div>
      </div>);
  }

  componentWillMount() {
    const SizeStyle = Quill.import('attributors/style/size');
    SizeStyle.whitelist = ['14px', '18px', '24px'];
    Quill.register(SizeStyle, true);

  }

  private _getWordCount = () : void => {
    const plainText = this.state.editorValue.replace(/<[^>]*>/g," ");
    const regex = /(?:\r\n|\r|\n)/g;  // new line, carriage return, line feed
    const cleanString = plainText.replace(regex, ' ').trim(); // replace above characters w/ space
    const wordArray = cleanString.match(/\S+/g);  // matches words according to whitespace
    this.setState({
      wordCount :wordArray ? wordArray.length : 0
    });
  }
}

export default QuillService;
