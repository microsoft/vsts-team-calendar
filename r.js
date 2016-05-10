#!/bin/sh
basedir=`dirname "$0"`

case `uname` in
    *CYGWIN*) basedir=`cygpath -w "$basedir"`;;
esac

if [ -x "$basedir/node" ]; then
  "$basedir/node"  "$basedir/node_modules/requirejs/bin/r.js" "$@"
  ret=$?
else 
  node  "$basedir/node_modules/requirejs/bin/r.js" "$@"
  ret=$?
fi
exit $ret
