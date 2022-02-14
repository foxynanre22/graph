// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
import React, { Component } from 'react';
import { BrowserRouter as Router, Route, Redirect } from 'react-router-dom';
import { Container } from 'reactstrap';
import withAuthProvider, { AuthComponentProps } from './common/AuthProvider';
import NavBar from './common//NavBar';
import ErrorMessage from './common/ErrorMessage';
import Welcome from './Welcome';
import Calendar from './components/Calendar/Calendar';
import NewEvent from './components/Calendar/NewEvent';
import 'bootstrap/dist/css/bootstrap.css';
import ChatTest from './components/Teams/ChatTest';

class App extends Component<AuthComponentProps> {
  render() {
    let error = null;
    if (this.props.error) {
      error = <ErrorMessage
        message={this.props.error.message}
        debug={this.props.error.debug} />;
    }

    // <renderSnippet>
    return (
      <Router>
        <div>
          <NavBar
            isAuthenticated={this.props.isAuthenticated}
            authButtonMethod={this.props.isAuthenticated ? this.props.logout : this.props.login}
            user={this.props.user} />
          <Container>
            {error}
            <Route exact path="/"
              render={(props) =>
                <Welcome {...props}
                  isAuthenticated={this.props.isAuthenticated}
                  user={this.props.user}
                  authButtonMethod={this.props.login} />
              } />
            <Route exact path="/calendar"
              render={(props) =>
                this.props.isAuthenticated ?
                  <Calendar {...props} /> :
                  <Redirect to="/" />
              } />
            <Route exact path="/newevent"
              render={(props) =>
                this.props.isAuthenticated ?
                  <NewEvent {...props} /> :
                  <Redirect to="/" />
              } />
            <Route exact path="/teams"
              render={(props) =>
                this.props.isAuthenticated ?
                  <ChatTest {...props} /> :
                  <Redirect to="/" />
              } />
          </Container>
        </div>
      </Router>
    );
    // </renderSnippet>
  }
}

export default withAuthProvider(App);
