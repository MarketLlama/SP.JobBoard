import * as React from 'react';
import ReactQuill, { Quill } from 'react-quill';
import styles from '../components/JobBoard.module.scss';
import 'react-quill/dist/quill.snow.css'; // ES6

export interface QuillServiceProps {
  onChange  : Function;
  defaultValue? : string;
  limit? : number;
}

export interface QuillServiceState {
 text : string;
 wordCount? : number;
}

class QuillService extends React.Component<QuillServiceProps, QuillServiceState> {
  constructor(props) {
    super(props)
    this.state = { text: '' }
  }


  modules = {
    toolbar: [
      [{ 'header': '1'}, {'header': '2'}, { 'font': [] }],
      [{ 'size': ['14px', '20px', '28px'] }],
      ['bold', 'italic', 'underline', 'strike', 'blockquote'],
      [{'list': 'ordered'}, {'list': 'bullet'},
       {'indent': '-1'}, {'indent': '+1'}],
      ['clean']
    ],
    clipboard: {
      // toggle to add extra line breaks when pasting HTML:
      matchVisual: true,
    }
  }

  formats = [
    'header', 'font', 'size',
    'bold', 'italic', 'underline', 'strike', 'blockquote',
    'list', 'bullet', 'indent'
  ]

  private _handleChange = (value) => {
    this.setState({ text: value })
    this.props.onChange(value);
    this._getWordCount();
    console.log(value);
  }

  render() {

    return (<div>
      <div >
        <ReactQuill value={this.state.text}
          modules={this.modules}
          formats={this.formats}
          onChange={this._handleChange}
        />
      </div>
        <div>
          <span>Word Count : {this.state.wordCount}</span>
        </div>
      </div>);
  }

  componentWillMount() {
    const SizeStyle = Quill.import('attributors/style/size');
    SizeStyle.whitelist = ['14px', '20px', '28px'];
    Quill.register(SizeStyle, true);

  }

  private _getWordCount = () : void => {
    const plainText = this.state.text.replace(/<[^>]*>/g," ");
    const regex = /(?:\r\n|\r|\n)/g;  // new line, carriage return, line feed
    const cleanString = plainText.replace(regex, ' ').trim(); // replace above characters w/ space
    const wordArray = cleanString.match(/\S+/g);  // matches words according to whitespace
    this.setState({
      wordCount :wordArray ? wordArray.length : 0
    });
  }
}

export default QuillService;
