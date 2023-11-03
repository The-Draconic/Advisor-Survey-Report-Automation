import dash
import dash_html_components as html

# Create a Dash web application
app = dash.Dash(__name__)

# Define the layout of the application
app.layout = html.Div(
    children=[
        html.H1("Hello World!"),
        html.P("This is a simple Dash web application."),
    ]
)

# Run the application
if __name__ == "__main__":
    app.run_server(debug=True)
