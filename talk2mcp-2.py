import os
from dotenv import load_dotenv
from mcp import ClientSession, StdioServerParameters, types
from mcp.client.stdio import stdio_client
import asyncio
from google import genai
from concurrent.futures import TimeoutError
from functools import partial

# Load environment variables from .env file
load_dotenv()

# Access your API key and initialize Gemini client correctly
api_key = os.getenv("GEMINI_API_KEY")
client = genai.Client(api_key=api_key)

max_iterations = 10
last_response = None
iteration = 0
iteration_response = []
email_sent = False

async def generate_with_timeout(client, prompt, timeout=10):
    """Generate content with a timeout"""
    print("Starting LLM generation...")
    try:
        # Convert the synchronous generate_content call to run in a thread
        loop = asyncio.get_event_loop()
        response = await asyncio.wait_for(
            loop.run_in_executor(
                None, 
                lambda: client.models.generate_content(
                    model="gemini-2.0-flash",
                    contents=prompt
                )
            ),
            timeout=timeout
        )
        print("LLM generation completed")
        return response
    except TimeoutError:
        print("LLM generation timed out!")
        raise
    except Exception as e:
        print(f"Error in LLM generation: {e}")
        raise

def reset_state():
    """Reset all global variables to their initial state"""
    global last_response, iteration, iteration_response, email_sent
    last_response = None
    iteration = 0
    iteration_response = []
    email_sent = False

async def main():
    max_retries = 3
    retry_count = 0
    
    while retry_count < max_retries:
        try:
            reset_state()  # Reset at the start of main
            print("Starting main execution...")
            
            # Create a single MCP server connection
            print("Establishing connection to MCP server...")
            server_params = StdioServerParameters(
                command="python",
                args=["example2-3.py", "--creds-file-path", "credentials.json", "--token-path", "token.json", "--dev"]
            )

            async with stdio_client(server_params) as (read, write):
                print("Connection established, creating session...")
                async with ClientSession(read, write) as session:
                    print("Session created, initializing...")
                    try:
                        await session.initialize()
                    except Exception as e:
                        print(f"Failed to initialize session: {e}")
                        continue
                    
                    # Get available tools
                    print("Requesting tool list...")
                    try:
                        tools_result = await session.list_tools()
                        tools = tools_result.tools
                        print(f"Successfully retrieved {len(tools)} tools")
                    except Exception as e:
                        print(f"Failed to get tool list: {e}")
                        continue
                    
                    # Create system prompt with available tools
                    print("Creating system prompt...")
                    print(f"Number of tools: {len(tools)}")
                    
                    try:
                        tools_description = []
                        for i, tool in enumerate(tools):
                            try:
                                params = tool.inputSchema
                                desc = getattr(tool, 'description', 'No description available')
                                name = getattr(tool, 'name', f'tool_{i}')
                                
                                if 'properties' in params:
                                    param_details = []
                                    for param_name, param_info in params['properties'].items():
                                        param_type = param_info.get('type', 'unknown')
                                        param_details.append(f"{param_name}: {param_type}")
                                    params_str = ', '.join(param_details)
                                else:
                                    params_str = 'no parameters'

                                tool_desc = f"{i+1}. {name}({params_str}) - {desc}"
                                tools_description.append(tool_desc)
                                print(f"Added description for tool: {tool_desc}")
                            except Exception as e:
                                print(f"Error processing tool {i}: {e}")
                                tools_description.append(f"{i+1}. Error processing tool")
                        
                        tools_description = "\n".join(tools_description)
                        print("Successfully created tools description")
                    except Exception as e:
                        print(f"Error creating tools description: {e}")
                        tools_description = "Error loading tools"
                    
                    print("Created system prompt...")
                    
                    # Define the system prompt
                    SYSTEM_PROMPT = """You are a structured reasoning assistant that helps solve mathematical problems, create presentations, and handle email communications. You must think through each step carefully and follow a strict format.

REASONING FRAMEWORK:
1. For each task, break it down into clear, sequential steps
2. Before each action, verify if you have all necessary information
3. After each step, validate the result before proceeding
4. If any step fails, identify the issue and try an alternative approach

RESPONSE FORMAT:
You must respond with EXACTLY ONE line in one of these formats:
1. For function calls:
   FUNCTION_CALL: function_name|param1|param2|...
2. For final answers:
   FINAL_ANSWER: [result]

AVAILABLE TOOLS AND USAGE:
{tools_description}

STEP-BY-STEP GUIDELINES:

1. Mathematical Operations:
   - Verify input types before calculation
   - Check for division by zero
   - Validate results are within expected ranges
   - Use appropriate mathematical functions (add, subtract, multiply, divide)

2. PowerPoint Creation:
   - Open presentation first: FUNCTION_CALL: open_powerpoint
   - Draw shapes: FUNCTION_CALL: draw_rectangle|x1|y1|x2|y2
   - Add text: FUNCTION_CALL: add_text_in_powerpoint|text
   - Close presentation: FUNCTION_CALL: close_powerpoint

3. Email Operations:
   - Required parameters for send_gmail:
     * recipient: email address
     * subject: email subject
     * body: email content
     * attachment_path: path to PowerPoint file
   - Format: FUNCTION_CALL: send_gmail|recipient|subject|body|attachment_path

ERROR HANDLING:
- If a function call fails, analyze the error and try an alternative approach
- For invalid inputs, request clarification
- If a tool is unavailable, explain the limitation and suggest alternatives

EXAMPLES:
1. Mathematical:
   FUNCTION_CALL: strings_to_chars_to_int|INDIA
   FUNCTION_CALL: int_list_to_exponential_sum|73,78,68,73,65

2. PowerPoint:
   FUNCTION_CALL: open_powerpoint
   FUNCTION_CALL: draw_rectangle|2|2|7|5
   FUNCTION_CALL: add_text_in_powerpoint|Final Result:\\n7.59982224609308e+33
   FUNCTION_CALL: close_powerpoint

3. Email:
   FUNCTION_CALL: send_gmail|padmanabhbosamia@gmail.com|ASCII Calculation Result|Input: INDIA\nASCII Values: [73, 78, 68, 73, 65]\nFinal Result: 7.59982224609308e+33|presentation.pptx

IMPORTANT:
- Think before responding
- Verify each step
- Follow the exact format
- No additional text or explanations
- One response per iteration
- Validate all inputs before use"""

                    query = """Find the ASCII values of characters in INDIA and then return sum of exponentials of those values. 
                    Create a PowerPoint presentation with the result in a rectangle, and then send an email to padmanabhbosamia@gmail.com with the results and attach the PowerPoint presentation.
                    
                    IMPORTANT: When sending the email with the send_gmail function, you MUST include the attachment_path parameter with the value 'presentation.pptx'.
                    The function call should look like this: FUNCTION_CALL: send_gmail|padmanabhbosamia@gmail.com|ASCII Calculation Result|Input: INDIA\nASCII Values: [73, 78, 68, 73, 65]\nFinal Result: 7.59982224609308e+33|presentation.pptx
                    
                    IMPORTANT: You must respond with EXACTLY ONE line in one of these formats:
                    1. For function calls: FUNCTION_CALL: function_name|param1|param2|...
                    2. For final answers: FINAL_ANSWER: [result]
                    
                    DO NOT include any explanations or additional text."""
                    print("Starting iteration loop...")
                    
                    # Use global iteration variables
                    global iteration, last_response, email_sent
                    
                    while iteration < max_iterations:
                        print(f"\n--- Iteration {iteration + 1} ---")
                        if last_response is None:
                            current_query = query
                        else:
                            current_query = current_query + "\n\n" + " ".join(iteration_response)
                            current_query = current_query + "  What should I do next?"

                        # Get model's response with timeout
                        print("Preparing to generate LLM response...")
                        prompt = f"{SYSTEM_PROMPT}\n\nQuery: {current_query}"
                        try:
                            response = await generate_with_timeout(client, prompt)
                            response_text = response.text.strip()
                            print(f"LLM Response: {response_text}")
                            
                            # Find the appropriate line in the response
                            for line in response_text.split('\n'):
                                line = line.strip()
                                if line.startswith(("FUNCTION_CALL:", "FINAL_ANSWER:")):
                                    response_text = line
                                    break
                            
                        except Exception as e:
                            print(f"Failed to get LLM response: {e}")
                            break

                        if response_text.startswith("FUNCTION_CALL:"):
                            _, function_info = response_text.split(":", 1)
                            parts = [p.strip() for p in function_info.split("|")]
                            func_name, params = parts[0], parts[1:]
                            
                            print(f"\nDEBUG: Raw function info: {function_info}")
                            print(f"DEBUG: Split parts: {parts}")
                            print(f"DEBUG: Function name: {func_name}")
                            print(f"DEBUG: Raw parameters: {params}")
                            
                            try:
                                # Find the matching tool to get its input schema
                                tool = next((t for t in tools if t.name == func_name), None)
                                if not tool:
                                    print(f"DEBUG: Available tools: {[t.name for t in tools]}")
                                    raise ValueError(f"Unknown tool: {func_name}")

                                print(f"DEBUG: Found tool: {tool.name}")
                                print(f"DEBUG: Tool schema: {tool.inputSchema}")

                                # Prepare arguments according to the tool's input schema
                                arguments = {}
                                schema_properties = tool.inputSchema.get('properties', {})
                                required_params = tool.inputSchema.get('required', [])
                                print(f"DEBUG: Schema properties: {schema_properties}")
                                print(f"DEBUG: Required parameters: {required_params}")

                                # Special handling for send_gmail function
                                if func_name == "send_gmail" and len(params) >= 4:
                                    # The fourth parameter is the attachment path
                                    attachment_path = params[3]
                                    print(f"DEBUG: Found attachment path: {attachment_path}")
                                    arguments["recipient"] = params[0]
                                    arguments["subject"] = params[1]
                                    arguments["body"] = params[2]
                                    arguments["attachment_path"] = attachment_path
                                else:
                                    # Normal parameter handling
                                    for param_name, param_info in schema_properties.items():
                                        # Skip optional parameters if we don't have enough values
                                        if param_name not in required_params and not params:
                                            print(f"DEBUG: Skipping optional parameter {param_name} as we don't have enough values")
                                            continue
                                            
                                        if not params:  # Check if we have enough parameters
                                            raise ValueError(f"Not enough parameters provided for {func_name}")
                                            
                                        value = params.pop(0)  # Get and remove the first parameter
                                        param_type = param_info.get('type', 'string')
                                        
                                        print(f"DEBUG: Converting parameter {param_name} with value {value} to type {param_type}")
                                        
                                        # Convert the value to the correct type based on the schema
                                        if param_type == 'integer':
                                            arguments[param_name] = int(value)
                                        elif param_type == 'number':
                                            arguments[param_name] = float(value)
                                        elif param_type == 'array':
                                            # Handle array input - if it's already a string representation of a list
                                            if value.startswith('[') and value.endswith(']'):
                                                # Parse the array string properly
                                                array_str = value.strip('[]')
                                                if array_str:
                                                    arguments[param_name] = [int(x.strip()) for x in array_str.split(',')]
                                                else:
                                                    arguments[param_name] = []
                                            else:
                                                # If it's a comma-separated string without brackets
                                                if ',' in value:
                                                    arguments[param_name] = [int(x.strip()) for x in value.split(',')]
                                                else:
                                                    # If it's a single value, make it a single-item list
                                                    arguments[param_name] = [int(value)]
                                        else:
                                            arguments[param_name] = str(value)

                                print(f"DEBUG: Final arguments: {arguments}")
                                print(f"DEBUG: Calling tool {func_name}")
                                
                                result = await session.call_tool(func_name, arguments=arguments)
                                print(f"DEBUG: Raw result: {result}")
                                
                                # Get the full result content
                                if hasattr(result, 'content'):
                                    print(f"DEBUG: Result has content attribute")
                                    # Handle multiple content items
                                    if isinstance(result.content, list):
                                        iteration_result = [
                                            item.text if hasattr(item, 'text') else str(item)
                                            for item in result.content
                                        ]
                                    else:
                                        iteration_result = str(result.content)
                                else:
                                    print(f"DEBUG: Result has no content attribute")
                                    iteration_result = str(result)
                                    
                                print(f"DEBUG: Final iteration result: {iteration_result}")
                                
                                # Format the response based on result type
                                if isinstance(iteration_result, list):
                                    result_str = f"[{', '.join(iteration_result)}]"
                                else:
                                    result_str = str(iteration_result)
                                
                                iteration_response.append(
                                    f"In the {iteration + 1} iteration you called {func_name} with {arguments} parameters, "
                                    f"and the function returned {result_str}."
                                )
                                last_response = iteration_result

                                # If this was the email sending operation, break after success
                                if func_name == "send_gmail" and "successfully" in result_str.lower():
                                    email_sent = True
                                    print("\n=== Agent Execution Complete ===")
                                    final_answer = next(resp.split("returned")[1].strip() 
                                        for resp in iteration_response 
                                        if "int_list_to_exponential_sum" in resp)
                                    print(f"Final Answer: {final_answer}")
                                    break

                            except Exception as e:
                                print(f"DEBUG: Error details: {str(e)}")
                                print(f"DEBUG: Error type: {type(e)}")
                                import traceback
                                traceback.print_exc()
                                iteration_response.append(f"Error in iteration {iteration + 1}: {str(e)}")
                                break

                        elif response_text.startswith("FINAL_ANSWER:"):
                            # Skip this section since we're handling final answer after email is sent
                            continue

                        iteration += 1

            break  # If we get here, everything worked fine
            
        except KeyboardInterrupt:
            print("\nKeyboard interrupt detected, cleaning up...")
            reset_state()
            break
        except Exception as e:
            print(f"Error in main execution (attempt {retry_count + 1}/{max_retries}): {e}")
            retry_count += 1
            if retry_count < max_retries:
                print(f"Retrying in 5 seconds...")
                await asyncio.sleep(5)
            else:
                print("Max retries reached, exiting...")
                raise
        finally:
            reset_state()  # Reset at the end of main

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nExiting due to keyboard interrupt...")
    except Exception as e:
        print(f"Fatal error: {e}")
        import traceback
        traceback.print_exc()
    
    
