FROM nginx:1.27.4-alpine

# Copy custom reverse proxy routing configuration
COPY nginx.conf /etc/nginx/conf.d/default.conf

# Expose public gateway port
EXPOSE 3000

# Run Nginx in foreground
CMD ["nginx", "-g", "daemon off;"]
