import resolve from 'rollup-plugin-node-resolve';
import commonjs from 'rollup-plugin-commonjs';

export default [{
  input: 'script.js',
  output: {
    file: 'bundle.js',
    format: 'iife'
  },
  name: 'bleh',
  plugins: [
    resolve(),
    commonjs()
  ]
}];